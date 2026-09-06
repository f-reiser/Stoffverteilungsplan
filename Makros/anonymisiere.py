#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt aus einer produktiven Mappe die anonyme Fassung fuer Vorlage/.

WARUM ES DIESE DATEI GIBT
    Die Mappen unter Vorlage/ sind die einzigen .xlsm im Repository. Sie
    muessen dort sein, damit die Dateiebenen (2, 3, 6, 8) in der CI
    ueberhaupt etwas zu pruefen haben - und die Mutationstests brauchen
    eine GEFUELLTE Mappe, in einer leeren laesst sich kein Fehler
    einbauen, der auffallen koennte.

    Gleichzeitig darf dort nichts Produktives stehen. Bis 06.09.2026 war
    Vorlage/Referenzmappe.xlsm eine byteweise Kopie des echten
    Mathematik-Gym-10-Plans: Themen, Termine, Schulaufgaben, Notizen,
    der Klarname der Lehrkraft, der Schulname und - im versteckten Namen
    wpPdfOrdner - der vollstaendige OneDrive-Pfad samt Windows-Benutzer.

    Deshalb ist Vorlage/ ein ERZEUGTES Verzeichnis. Es wird nicht von
    Hand gepflegt, sondern neu gebaut. Damit kann es auch nicht mehr
    von der Arbeitsdatei abweichen, ohne dass es jemand merkt.

AUFRUF
    python3 anonymisiere.py ../Stoffverteilungsplan_Template.xlsm \\
                            ../Vorlage/Stoffverteilungsplan_Template.xlsm

    Danach IMMER pruefen:
        python3 pruefe_anonym.py ../Vorlage/*.xlsm
        python3 pruefe_alles.py

VORGEHEN
    Ausschliesslich XML-Chirurgie im entpackten Zip. openpyxl darf diese
    Mappen NIE speichern - es wirft die x14-Erweiterungen weg und
    beschaedigt xl/vbaProject.bin (Projektregel 3).

    Angefasst wird nur xl/sharedStrings.xml, und dort nur die Eintraege,
    die AUSSCHLIESSLICH in Inhaltszellen stehen. Alles, was auch nur
    einmal in einer Kopfzeile, einer Beschriftung oder einer Liste
    auftaucht, bleibt unangetastet - "Summe", "nein" (Fixierung),
    "Schulliste" und die Kategorienliste M2:M6 treiben Selbsttest,
    Gueltigkeitspruefung und bedingte Formatierung.
"""
import os
import re
import shutil
import sys
import zipfile

BLATT = {"sheet1.xml": "Steuerung", "sheet2.xml": "Wochenplan",
         "sheet3.xml": "Lernbereiche", "sheet4.xml": "Einstellungen",
         "sheet5.xml": "Anleitung"}

#  Klarnamen: unabhaengig davon, wo sie stehen.
PERSONEN = {
    "Florian Reiser": "Max Mustermann",
    "Schule 1": "Beispiel-Gymnasium",
    "Schule 2": "Beispiel-FOS",
}

#  Inhaltsspalten je Blatt (Datenzeilen, nicht Kopf, nicht Summenzeile).
INHALT = {
    ("Wochenplan", "G"): "Thema",
    ("Wochenplan", "H"): "Inhalt",
    ("Wochenplan", "I"): "Material",
    ("Wochenplan", "M"): "Notiz",
    ("Lernbereiche", "D"): "Lernbereich",
    ("Lernbereiche", "J"): "Kompetenz",
}


def zellen(x):
    for ref, idx in re.findall(
            r'<c r="([A-Z]+\d+)"[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', x):
        yield (re.match(r"[A-Z]+", ref).group(0),
               int(re.search(r"\d+", ref).group(0)), int(idx))


def anonymisiere(quelle, ziel):
    z = zipfile.ZipFile(quelle)
    sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    si = re.findall(r"<si>.*?</si>", sx, re.S)

    def text(s):
        return "".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S))

    # --- Wo wird welcher Eintrag benutzt? Ueber ALLE Blaetter. --------
    nutzung = {}
    summenzeile = {}
    for datei, name in BLATT.items():
        pfad = "xl/worksheets/" + datei
        if pfad not in z.namelist():
            continue
        x = z.read(pfad).decode("utf-8")
        for sp, nr, idx in zellen(x):
            nutzung.setdefault(idx, set()).add((name, sp, nr))
            if name == "Lernbereiche" and sp == "D" and text(si[idx]) == "Summe":
                summenzeile[name] = nr

    # --- Welche Eintraege duerfen ersetzt werden? --------------------
    ersatz = {}
    lb_nr = 0
    for idx in sorted(nutzung):
        alt = text(si[idx])
        if alt in PERSONEN:
            ersatz[idx] = PERSONEN[alt]
            continue
        orte = nutzung[idx]
        #  Nur wenn JEDE Verwendung eine Inhaltszelle ist.
        arten = set()
        for name, sp, nr in orte:
            if nr <= 3 or nr == summenzeile.get(name):
                arten.add(None)
            else:
                arten.add(INHALT.get((name, sp)))
        if None in arten or not arten or arten == {None}:
            continue
        if "Lernbereich" in arten:
            lb_nr += 1
            ersatz[idx] = "Lernbereich %d (Platzhalter)" % lb_nr
        elif "Kompetenz" in arten:
            ersatz[idx] = "Kompetenzerwartungen (Platzhalter)"
        elif "Thema" in arten:
            ersatz[idx] = "Thema (Platzhalter)"
        elif "Inhalt" in arten:
            ersatz[idx] = "Inhalt der Stunde (Platzhalter)"
        elif "Material" in arten:
            ersatz[idx] = "Material (Platzhalter)"
        elif "Notiz" in arten:
            ersatz[idx] = "Notiz (Platzhalter)"

    # --- Neue sharedStrings bauen ------------------------------------
    neu = sx
    for idx, wert in ersatz.items():
        alt_si = si[idx]
        neu_si = '<si><t xml:space="preserve">%s</t></si>' % wert
        assert alt_si in neu, "Eintrag %d nicht gefunden" % idx
        neu = neu.replace(alt_si, neu_si, 1)

    # --- Zweiter Durchgang: ueber ALLE XML-Teile ---------------------
    #  Klarnamen stecken nicht nur in Inhaltszellen:
    #    - im Titelblock (Zeile 1/2), den die Kopfzeilenregel oben
    #      absichtlich schuetzt,
    #    - in docProps/core.xml als lastModifiedBy,
    #    - im versteckten Namen wpPdfOrdner, der den vollstaendigen
    #      OneDrive-Pfad samt Schulname und Windows-Benutzernamen fuehrt.
    #  Der wird ersatzlos entfernt: modKopf merkt sich damit nur den
    #  zuletzt benutzten Export-Ordner und kommt ohne ihn aus
    #  (GemerkterOrdner liefert dann "").
    teile = {"xl/sharedStrings.xml": neu}
    for n in z.namelist():
        if not n.endswith(".xml") and not n.endswith(".rels"):
            continue
        t = teile.get(n)
        if t is None:
            t = z.read(n).decode("utf-8")
        vorher = t
        t = re.sub(r'<definedName name="wpPdfOrdner".*?</definedName>', "", t)
        for alt, wert in PERSONEN.items():
            t = t.replace(alt, wert)
        if t != vorher:
            teile[n] = t

    # --- Neu packen: [Content_Types].xml zuerst ----------------------
    namen = z.namelist()
    reihe = (["[Content_Types].xml"]
             + [n for n in namen if n != "[Content_Types].xml"])
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            daten = (teile[n].encode("utf-8") if n in teile else z.read(n))
            out.writestr(z.getinfo(n), daten)
    z.close()
    return len(ersatz), sorted(teile)


if __name__ == "__main__":
    quelle, ziel = sys.argv[1], sys.argv[2]
    n, teile = anonymisiere(quelle, ziel)
    print("%s\n   %d Texteintraege ersetzt, geaenderte Teile: %s"
          % (os.path.basename(quelle), n, ", ".join(teile)))
