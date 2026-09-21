#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bringt die VBA-Module aus Makros/ in eine .xlsm - per Excel-COM, nicht per
XML-Chirurgie (Issue #77).

WARUM ES DIESES SKRIPT GIBT
    xl/vbaProject.bin ist ein kompiliertes OLE-Compound-File, kein Text -
    anders als sharedStrings.xml (anonymisiere.py) laesst es sich nicht per
    Zip/XML-Bearbeitung aus Quelltext bauen. Nur Excel selbst kompiliert VBA.
    Deshalb automatisiert dieses Skript genau das, was bisher Alt+F11 und
    Handimport waren: ueber das VBA-Projektobjektmodell (braucht "Zugriff auf
    das VBA-Projektobjektmodell vertrauen" im Trust Center von Excel).

    Laeuft NUR lokal, nie in der CI - dort steht kein Excel (siehe Ebene 4/5,
    Doku/Testebenen.md). Dieselbe Grenze wie beim Excel-Selbsttest.

WAS EINGESETZT WIRD
    Alle mod*.bas ausser modKonfig.bas - das Kennwort wird NIE ausgeliefert
    (Projektregel 4). An seiner Stelle geht modKonfig.bas.vorlage hinein,
    unter dem Namen modKonfig.
    Dazu die zwei Dokumentmodule DieseArbeitsmappe_Modul.txt (Modul
    "ThisWorkbook") und Wochenplan_Blattmodul.txt (Blattmodul der Tabelle,
    die als "Wochenplan" angezeigt wird - das CodeName wird zur Laufzeit
    ueber den Blattnamen aufgeloest, nicht geraten).

AUFRUF
    python3 makros_einsetzen.py --pruefen <mappe.xlsm>
        Nur lesen: vergleicht die Module IN der Mappe mit dem Stand in
        Makros/. Exit 0 = identisch, Exit 1 = Abweichung (mit Liste).

    python3 makros_einsetzen.py <mappe.xlsm>
        Setzt die Module ein und speichert - mit Excel selbst, nie mit
        openpyxl (Projektregel 3).

    python3 makros_einsetzen.py --referenzmappe <template.xlsm> <alte_referenzmappe.xlsm> <ziel.xlsm>
        Baut eine Referenzmappe aus dem (schon aktualisierten) Template neu -
        ueber modUebernahme, denselben Weg wie beim echten Update. Nimmt
        NIE die befuellte produktive Mappe, sondern die alte Referenzmappe
        selbst als Datenquelle (Issue #77).

VORHER SPERREN
    Vorlage/*.xlsm liegen per Git LFS mit Sperre (CLAUDE.md). Dieses Skript
    sperrt selbst nicht - das gehoert an den Aufrufer (git-branch-strategie).
"""
import glob
import io
import os
import re
import shutil
import sys
import tempfile
import zipfile

HIER = os.path.dirname(os.path.abspath(__file__))

#  vbext_ct_StdModule - als Zahl, nicht ueber win32com.client.constants:
#  Letzteres braucht eine generierte Typelib (EnsureDispatch/makepy), Ersteres
#  nicht. So laeuft das Skript auch dann, wenn "python -m win32com.client.makepy"
#  nie aufgerufen wurde.
VBEXT_CT_STDMODULE = 1

DOKUMENT_MODULE = {
    "DieseArbeitsmappe_Modul.txt": "ThisWorkbook",
    "Wochenplan_Blattmodul.txt": "Wochenplan",   #  Blattname, nicht CodeName
}

_ATTRIBUTE_ZEILE = re.compile(r"^Attribute VB_\w+ = .*\r?\n?", re.MULTILINE)


def cp1252_text(pfad):
    """Liest eine .bas/.txt-Datei so, wie der VBA-Editor sie versteht.

    io.open(..., encoding='cp1252', newline='') - Projektregel 1. NIE ueber
    Edit/Write anfassen, dieselbe Regel gilt sinngemaess fuers Lesen: ein
    falscher Codec liest Umlaute falsch, noch bevor irgendwas geschrieben
    wird.
    """
    with io.open(pfad, encoding="cp1252", newline="") as f:
        return f.read()


def ohne_attribute(text):
    """CodeModule.Lines() aus einem LAUFENDEN VBA-Projekt enthaelt die
    "Attribute VB_..."-Zeilen nicht - die verwaltet Excel separat und
    schreibt sie nur beim Export in die Datei. Fuer den Vergleich muessen
    sie deshalb auch aus der Quelldatei heraus."""
    return _ATTRIBUTE_ZEILE.sub("", text, count=8)


def normalisiert(text):
    """CodeModule.Lines() liefert IMMER CRLF, unabhaengig davon, mit welchen
    Zeilenenden ein Modul importiert oder eingefuegt wurde - eine VBA-
    Eigenheit, keine inhaltliche Abweichung. Ohne diese Angleichung waere
    jeder Vergleich blind rot, selbst bei identischem Inhalt (das genau ist
    beim ersten Lauf gegen die echte Mappe aufgefallen)."""
    return text.replace("\r\n", "\n")


def standard_module_verzeichnis():
    """Arbeitsverzeichnis mit allen mod*.bas AUSSER modKonfig.bas - an
    dessen Stelle IMMER modKonfig.bas.vorlage, nie das echte Modul
    (Projektregel 4). Anders als pruefe_alles.py.module_bereitstellen(),
    das fuer die lokale Pruefung das echte Modul bevorzugt: was hier
    entsteht, geht in eine ausgelieferte Datei.

    Rueckgabe: (Verzeichnis, sortierte Liste der .bas-Pfade, Aufraeumfunktion)
    """
    vorlage = os.path.join(HIER, "modKonfig.bas.vorlage")
    if not os.path.exists(vorlage):
        raise SystemExit("modKonfig.bas.vorlage fehlt - ohne sie kann "
                          "modKonfig nicht ausgeliefert werden.")

    echte = [p for p in sorted(glob.glob(os.path.join(HIER, "mod*.bas")))
             if os.path.basename(p) != "modKonfig.bas"]

    tmp = tempfile.mkdtemp(prefix="makros_einsetzen_")
    for p in echte:
        shutil.copy2(p, tmp)
    ziel_konfig = os.path.join(tmp, "modKonfig.bas")
    shutil.copy2(vorlage, ziel_konfig)
    return tmp, sorted(glob.glob(os.path.join(tmp, "mod*.bas"))), \
        lambda: shutil.rmtree(tmp, ignore_errors=True)


def modulname(bas_pfad):
    """Attribute VB_Name aus einer .bas-Datei - unabhaengig vom Dateinamen,
    genau wie Ebene 7 (pruefe_module.py) es beim Ausliefern verlangt."""
    m = re.search(r'^Attribute VB_Name = "([^"]+)"',
                  cp1252_text(bas_pfad), re.MULTILINE)
    if not m:
        raise SystemExit("%s: kein Attribute VB_Name gefunden." % bas_pfad)
    return m.group(1)


def soll_stand():
    """Der komplette Soll-Zustand: {Modulname: Quelltext ohne Attribute-
    Zeilen}. Fasst Standardmodule und Dokumentmodule zusammen - fuer den
    Vergleich (--pruefen) zaehlt nur, was am Ende im CodeModule steht,
    nicht WIE es dort hineinkam."""
    tmp, dateien, aufraeumen = standard_module_verzeichnis()
    try:
        stand = {modulname(p): normalisiert(ohne_attribute(cp1252_text(p)))
                  for p in dateien}
    finally:
        aufraeumen()

    for datei, ziel in DOKUMENT_MODULE.items():
        stand[ziel] = normalisiert(cp1252_text(os.path.join(HIER, datei)))
    return stand


#  ------------------------------------------------------------- Excel-COM
#  Erst ab hier wird win32com gebraucht - lazy importiert, damit soll_stand()
#  und die Vergleichslogik auch ohne pywin32/Excel testbar bleiben.

def _oeffnen(mappe_pfad):
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(mappe_pfad))
    except Exception:
        xl.Quit()
        raise
    return xl, wb


def _dokument_codenamen(wb):
    """{tatsaechlicher VBA-Komponentenname: unser logischer Name}.

    wb.CodeName ist die stabile, sprachunabhaengige Kennung des
    Arbeitsmappen-Moduls - NICHT der String "ThisWorkbook": In dieser
    Mappe heisst die Komponente wirklich "DieseArbeitsmappe" (sichtbar in
    VBComponents, nicht nur im VBA-Editor). Eine feste Zeichenkette waere
    hier ein Fehler gewesen, keine Abkuerzung.
    """
    return {wb.CodeName: "ThisWorkbook",
            wb.Worksheets("Wochenplan").CodeName: "Wochenplan"}


def _ist_stand(wb):
    """Liest den aktuellen Modulinhalt aus einer offenen Mappe - dieselbe
    Form wie soll_stand(), fuer den direkten Vergleich."""
    doc_namen = _dokument_codenamen(wb)
    stand = {}
    for comp in wb.VBProject.VBComponents:
        cm = comp.CodeModule
        text = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines > 0 else ""
        stand[doc_namen.get(comp.Name, comp.Name)] = normalisiert(text)
    return stand


def pruefen(mappe_pfad, soll=None):
    """--pruefen: vergleicht Ist- und Soll-Stand, druckt Abweichungen.
    Rueckgabe: 0 wenn identisch, 1 sonst - als Exit-Code verwendbar."""
    soll = soll_stand() if soll is None else soll
    xl, wb = _oeffnen(mappe_pfad)
    try:
        ist = _ist_stand(wb)
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()

    fehler = []
    for name, text in sorted(soll.items()):
        if name not in ist:
            fehler.append("%s: fehlt in der Mappe" % name)
        #  .lower(): VBA vereinheitlicht die Gross-/Kleinschreibung eines
        #  Bezeichners projektweit auf die ERSTE gefundene Schreibweise -
        #  unabhaengig vom semantischen Zusammenhang. modSelbsttest.bas
        #  deklariert einen Parameter "name" (klein); dadurch erscheint
        #  jedes spaetere "Shapes(i).Name" im ganzen Projekt als "...name".
        #  Das passiert genauso bei einem manuellen Alt+F11-Import und ist
        #  keine inhaltliche Abweichung, die dieser Vergleich melden soll.
        #  .rstrip("\n"): CodeModule.Lines() liefert nie eine abschliessende
        #  Leerzeile nach der letzten Codezeile, jede .bas-Datei aber schon
        #  (Ebene 7 verlangt genau das). Auch das ist Form, kein Inhalt.
        elif ist[name].lower().rstrip("\n") != text.lower().rstrip("\n"):
            fehler.append("%s: Inhalt weicht ab" % name)
    if fehler:
        print("Abweichungen zwischen Mappe und Makros/:")
        for f in fehler:
            print("  " + f)
        return 1
    print("Mappe stimmt mit Makros/ ueberein: %d Module geprueft." % len(soll))
    return 0


def einsetzen(mappe_pfad):
    """Setzt alle Module aus Makros/ in die Mappe ein und speichert - mit
    Excel selbst. Standardmodule: komplett entfernen und neu importieren
    (Import() liest die Datei direkt, keine Python-seitige Dekodierung
    noetig). Dokumentmodule: Code-Zeilen ersetzen, das Modul selbst gibt es
    in jeder Mappe schon.

    Vor dem Austausch wird entschuetzt, danach neu geschuetzt (ueber
    Setup_Stoffverteilungsplan, das intern Blattschutz_Einrichten aufruft) -
    beides mit dem JEWEILS aktuellen SCHUTZ_PW. Ohne das bliebe der
    Blattschutz auf dem alten Kennwort haengen: Der Blattschutz selbst wird
    mit der Datei gespeichert (Hash in der Worksheet-XML), das ausgetauschte
    modKonfig traegt aber sofort ein anderes Kennwort (die Vorlage) - beide
    liefen sonst auseinander, und "Einrichtung / Reparatur" koennte die
    eigenen Blaetter nicht mehr entsperren. Derselbe Fallstrick hat den
    ersten Versuch, daraus eine Referenzmappe abzuleiten, mit "Blattschutz
    liess sich nicht aufheben" scheitern lassen."""
    tmp, dateien, aufraeumen = standard_module_verzeichnis()
    try:
        xl, wb = _oeffnen(mappe_pfad)
        try:
            app = xl.Application
            vbproj = wb.VBProject

            app.Run("modSchutz.Schutz_Aus")

            fuer_entfernen = [c for c in vbproj.VBComponents
                               if c.Type == VBEXT_CT_STDMODULE]
            for comp in fuer_entfernen:
                vbproj.VBComponents.Remove(comp)
            for pfad in dateien:
                vbproj.VBComponents.Import(pfad)

            codename_von = {logisch: codename for codename, logisch
                             in _dokument_codenamen(wb).items()}
            for datei, logischer_name in DOKUMENT_MODULE.items():
                comp = vbproj.VBComponents(codename_von[logischer_name])
                cm = comp.CodeModule
                if cm.CountOfLines > 0:
                    cm.DeleteLines(1, cm.CountOfLines)
                cm.AddFromString(cp1252_text(os.path.join(HIER, datei)))

            app.Run("modWochenplan.ClearLastError")
            app.Run("modWochenplan.SetQuiet", True)
            try:
                app.Run("modSteuerung.Setup_Stoffverteilungsplan")
            finally:
                app.Run("modWochenplan.SetQuiet", False)
            fehler = app.Run("modWochenplan.LastError")
            if fehler:
                raise RuntimeError("Einrichtung nach dem Modulaustausch "
                                    "fehlgeschlagen: " + fehler)

            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
            xl.Quit()
    finally:
        aufraeumen()
    saeubere_metadaten(mappe_pfad)
    print("Makros eingesetzt: " + os.path.abspath(mappe_pfad))


_SHEET_PROTECTION = re.compile(rb"<sheetProtection\b.*?/>", re.S)

#  Dieselben zwei Spuren, die anonymisiere.py im "zweiten Durchgang" entfernt
#  (dort fuer eine produktive Quelle). Hier noetig, weil JEDES Speichern
#  durch Excel - auch unseres per COM - den aktuellen Windows-Benutzer neu
#  hineinschreibt, unabhaengig vom Zellinhalt: lastModifiedBy in
#  docProps/core.xml und der Pfad in x15ac:absPath (workbook.xml). Ohne
#  dieses Nachfassen meldet pruefe_anonym.py (Ebene A) nach jedem Lauf
#  dieses Skripts einen Fund (beim ersten Referenzmappe-Neubau geschehen).
_ABSPATH_BLOCK = re.compile(
    rb"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*?"
    rb"absPath.*?</mc:AlternateContent>", re.S)
_WPPDFORDNER = re.compile(rb'<definedName name="wpPdfOrdner".*?</definedName>', re.S)


def saeubere_metadaten(pfad):
    """Entfernt Windows-Benutzerpfad und Klarname, die Excel bei JEDEM
    Speichern selbst hinterlaesst - unabhaengig vom Inhalt der Zellen.
    Arbeitet in-place, atomar (schreiben unter neuem Namen, dann ersetzen -
    wie anonymisiere.py und logos_einsetzen.py)."""
    import pruefe_anonym
    personen = pruefe_anonym.lokale_datei()["ersetzungen"]

    z = zipfile.ZipFile(pfad)
    reihe = (["[Content_Types].xml"]
             + [n for n in z.namelist() if n != "[Content_Types].xml"])
    tmp = pfad + ".neu"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            daten = z.read(n)
            if n.endswith(".xml") or n.endswith(".rels"):
                daten = _WPPDFORDNER.sub(b"", daten)
                daten = _ABSPATH_BLOCK.sub(b"", daten)
                for alt, wert in personen.items():
                    daten = daten.replace(alt.encode("utf-8"), wert.encode("utf-8"))
            out.writestr(z.getinfo(n), daten)
    z.close()
    os.replace(tmp, pfad)


def entschuetzte_kopie(quelle_pfad, ziel_pfad):
    """Kopie einer .xlsm OHNE Blattschutz - reine XML-Chirurgie im
    entpackten Zip (Projektregel 3), keine VBA-Beruehrung.

    Wozu: modUebernahme muss die ALTE Mappe lesen koennen, auch wenn ihr
    Blattschutz mit dem ECHTEN SCHUTZ_PW gesetzt wurde. Das Kennwort selbst
    darf dieses Skript nie kennen oder brauchen (Projektregel 4) - also wird
    die Sperre entfernt, statt sie mit einem Kennwort aufzuheben.
    """
    z = zipfile.ZipFile(quelle_pfad)
    reihe = (["[Content_Types].xml"]
             + [n for n in z.namelist() if n != "[Content_Types].xml"])
    with zipfile.ZipFile(ziel_pfad, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            daten = z.read(n)
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                daten = _SHEET_PROTECTION.sub(b"", daten)
            out.writestr(z.getinfo(n), daten)
    z.close()


def referenzmappe_neu_bauen(template_pfad, alte_referenzmappe_pfad, ziel_pfad):
    """Baut Referenzmappe.xlsm aus dem (bereits aktualisierten) Template neu -
    per modUebernahme, demselben Weg, den auch ein Kollege beim Update geht.
    Nie aus der befuellten produktiven Mappe (die nicht im Repository liegt
    und nicht anonymisiert werden soll) - Issue #77, Kommentar vom
    21.09.2026: der Weg ueber "erst befuellen, dann anonymisieren" waere ein
    unnoetiger Umweg und teste nicht den echten Update-Pfad mit.

    Wirft RuntimeError mit der Meldung aus modWochenplan.LastError(), wenn
    die Uebernahme fehlschlaegt - z.B. bei einer Vorlage, deren Layout nicht
    mehr zu den Beschriftungen passt, an denen modUebernahme sich orientiert.
    """
    arbeitsverzeichnis = tempfile.mkdtemp(prefix="referenzmappe_neu_")
    try:
        kopie_pfad = os.path.join(arbeitsverzeichnis,
                                   os.path.basename(ziel_pfad))
        shutil.copy2(template_pfad, kopie_pfad)

        entschuetzte_alte_pfad = os.path.join(arbeitsverzeichnis,
                                               "alte_ohne_schutz.xlsm")
        entschuetzte_kopie(alte_referenzmappe_pfad, entschuetzte_alte_pfad)

        xl, wb = _oeffnen(kopie_pfad)
        try:
            app = xl.Application
            app.Run("modWochenplan.ClearLastError")
            app.Run("modWochenplan.SetQuiet", True)
            try:
                nPlan = app.Run("modUebernahme.UebernahmeAusfuehren",
                                 entschuetzte_alte_pfad, "", 0, 0, 0)
                if nPlan >= 0:
                    app.Run("modSteuerung.Setup_Stoffverteilungsplan")
                    app.Run("modKalender.UW_Und_Ferien_Generieren")
            finally:
                app.Run("modWochenplan.SetQuiet", False)

            fehler = app.Run("modWochenplan.LastError")
            if nPlan < 0 or fehler:
                raise RuntimeError("Referenzmappe-Neubau fehlgeschlagen: "
                                    + (fehler or "UebernahmeAusfuehren meldet -1, "
                                       "aber LastError ist leer."))
            wb.SaveAs(os.path.abspath(ziel_pfad),
                      FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        finally:
            wb.Close(SaveChanges=False)
            xl.Quit()
    finally:
        shutil.rmtree(arbeitsverzeichnis, ignore_errors=True)
    saeubere_metadaten(ziel_pfad)
    print("Referenzmappe neu gebaut: %d Planzeilen, Ziel: %s"
          % (nPlan, os.path.abspath(ziel_pfad)))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    if sys.argv[1] == "--pruefen":
        if len(sys.argv) != 3:
            raise SystemExit(__doc__)
        raise SystemExit(pruefen(sys.argv[2]))
    elif sys.argv[1] == "--referenzmappe":
        if len(sys.argv) != 5:
            raise SystemExit(__doc__)
        referenzmappe_neu_bauen(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        einsetzen(sys.argv[1])
