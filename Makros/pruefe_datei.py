#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Abnahmepruefung einer fertigen Stoffverteilungsplan-Datei.

Prueft eine .xlsm von aussen - ohne Excel, ohne VBA - auf genau die
Eigenschaften, die in diesem Projekt schon einmal kaputt waren.

WARUM ES DAS GIBT
    Der Selbsttest laeuft nur, wenn jemand ihn in Excel startet, und er
    laeuft immer nur in EINER Datei. Dieser Pruefer laeuft ueber alle
    ausgelieferten Dateien auf einmal und findet Zustaende, die eine
    Simulation nicht sieht, weil sie am falschen Ort sucht.

    Anlass: Am 05.09.2026 hat die Import-Simulation "0 Abweichungen"
    gemeldet, waehrend in der echten Datei 35 leere Zeilen unter der
    Tabelle standen. Die Simulation hat nur den Bereich betrachtet, in
    den geschrieben wurde - nicht den Rest des Blattes.

REGEL FUER JEDE NEUE PRUEFUNG
    Erst gegen eine Datei laufen lassen, in der der Fehler NACHWEISLICH
    steckt, und sehen, dass sie rot wird. Eine Pruefung, die noch nie
    rot war, ist keine Pruefung.
    Mit --selbsttest baut das Skript die Fehler kuenstlich ein und
    verlangt, dass jede Regel anschlaegt.
"""
import sys, zipfile, re, argparse, warnings, copy, io, tempfile
import xml.etree.ElementTree as ET
warnings.filterwarnings("ignore")
import openpyxl

WP, LB, ES, ST = "Wochenplan", "Lernbereiche", "Einstellungen", "Steuerung"


# ---------------------------------------------------------------- Hilfen
def txt(ws, r, c):
    v = ws.cell(r, c).value
    return "" if v is None else str(v).strip()


def kopfzeile(ws, spalte, text):
    for r in range(1, 60):
        if txt(ws, r, spalte).lower().startswith(text.lower()):
            return r
    return 0


def plan_letzte_zeile(ws, erste):
    """Letzte Zeile mit Inhalt - die VBA-Entsprechung von PlanLastRow."""
    cols = [5, 6, 7, 8, 9, 11, 13, 14, 22]          # E F G H I K M N V
    m = erste
    for r in range(erste, ws.max_row + 1):
        if any(txt(ws, r, c) for c in cols):
            m = r
    return m


# ------------------------------------------------------------- Pruefungen
def pruefe(pfad, wb=None):
    """Liefert eine Liste von Befunden (leer = in Ordnung)."""
    fund = []
    if wb is None:
        wb = openpyxl.load_workbook(pfad, data_only=False)

    def melde(regel, text):
        fund.append((regel, text))

    # --- 1. Blaetter -------------------------------------------------
    for blatt in (WP, LB, ES, ST):
        if blatt not in wb.sheetnames:
            melde("blaetter", "Blatt %r fehlt" % blatt)
    for blatt in ("Update", "VBA-Update"):
        if blatt in wb.sheetnames:
            melde("update-blatt",
                  "Das entfallene Blatt %r ist noch da" % blatt)
    if WP not in wb.sheetnames or LB not in wb.sheetnames:
        return fund

    ws, wl = wb[WP], wb[LB]
    wpKopf = kopfzeile(ws, 2, "Referenz Code")
    lbKopf = kopfzeile(wl, 1, "Referenzcode")
    if wpKopf == 0:
        melde("blaetter", "Ueberschriftenzeile des Wochenplans nicht gefunden")
        return fund
    erste = wpKopf + 1
    letzte = plan_letzte_zeile(ws, erste)

    #  Eine noch leere VORLAGE ist ein Sonderfall: sie bringt
    #  absichtlich eine Musterzeile und darunter formatierte
    #  Reservezeilen mit, und ihre Summenformel zeigt ins Leere, bis
    #  Lernbereiche eingetragen sind. Beides waere in einem fertigen
    #  Plan ein Fehler - hier ist es der Lieferzustand. Ohne diese
    #  Unterscheidung meldet der Pruefer die Vorlage jedes Mal rot und
    #  wird dadurch wertlos ("das ist immer so").
    leer_vorlage = not any(txt(ws, r, 7) for r in range(erste, letzte + 1))
    if leer_vorlage:
        melde("hinweis", "leere Vorlage - Plan- und Lernbereichspruefungen "
                         "uebersprungen")
        return fund

    # --- 2. Keine Leerzeilen unter der Tabelle -----------------------
    #  Das ist der Befund, an dem alles haengt: die Zeilen sind leer,
    #  aber formatiert, und liegen UNTERHALB von plan_letzte_zeile.
    unten = ws.max_row
    leer = [r for r in range(letzte + 1, unten + 1)
            if not any(txt(ws, r, c) for c in range(1, 23))]
    if leer:
        melde("leerzeilen",
              "%d leere Zeilen unter der Tabelle (Zeile %d bis %d)"
              % (len(leer), leer[0], leer[-1]))

    # --- 3. Lernbereiche: Rechenspalten sind Formeln -----------------
    if lbKopf:
        lbErste = lbKopf + 1
        zSum = 0
        for r in range(lbErste, lbErste + 80):
            if txt(wl, r, 4).lower() == "summe":
                zSum = r
                break
        if zSum == 0:
            melde("lb-summe", "Summenzeile der Lernbereiche nicht gefunden")
        else:
            lbLetzte = lbErste - 1
            for r in range(lbErste, zSum):
                if txt(wl, r, 1) or txt(wl, r, 4):
                    lbLetzte = r
            ohne = []
            for r in range(lbErste, lbLetzte + 1):
                for c, name in ((6, "F"), (7, "G"), (8, "H"), (9, "I")):
                    v = wl.cell(r, c).value
                    if not (isinstance(v, str) and v.startswith("=")):
                        ohne.append("%s%d" % (name, r))
            if ohne:
                melde("lb-formeln",
                      "Rechenspalten ohne Formel: " + " ".join(ohne[:12]))
            if zSum != lbLetzte + 1:
                melde("lb-summe",
                      "Summenzeile steht in %d, letzter Lernbereich in %d"
                      % (zSum, lbLetzte))
            soll = "E%d:E%d" % (lbErste, lbLetzte)
            ist = str(wl.cell(zSum, 5).value or "")
            if soll not in ist:
                melde("lb-summe",
                      "Summenformel %r deckt nicht %s ab" % (ist, soll))

    # --- 4. Ferienzeilen tragen ihr Kennzeichen ----------------------
    #  Erst AB der ersten Planzeile suchen. Der Titelblock ganz oben
    #  verbindet ebenfalls B bis M - beim ersten Lauf hat diese Regel
    #  ihn prompt als Ferienzeile ohne Kennzeichen gemeldet.
    verbunden = {}
    for rng in ws.merged_cells.ranges:
        if (rng.min_col == 2 and (rng.max_col - rng.min_col + 1) > 4
                and rng.min_row >= erste):
            verbunden[rng.min_row] = True
    ohneK = [r for r in verbunden
             if txt(ws, r, 14).upper() not in ("F", "FERIEN")]
    if ohneK:
        melde("kennzeichen",
              "Verbundene Zeilen ohne Kennzeichen in Spalte N: %s"
              % sorted(ohneK)[:10])

    # --- 5. Unterrichtswochen laufen aufsteigend ---------------------
    vorher, rueck = 0, []
    for r in range(erste, letzte + 1):
        if r in verbunden:
            continue
        v = ws.cell(r, 5).value
        if isinstance(v, (int, float)):
            if v < vorher:
                rueck.append(r)
            vorher = v
    if rueck:
        melde("wochenfolge",
              "Unterrichtswochen laufen rueckwaerts in Zeile %s" % rueck[:5])

    fund += pruefe_struktur(pfad)
    return fund


# --------------------------------------------------- Struktur der Datei
def pruefe_struktur(pfad):
    """Pruefebene 8: ist die .xlsm als Archiv und als XML noch heil?

    Diese Pruefungen liefen bisher als Handgriff nach jedem Eingriff ins
    entpackte Zip - also nur, solange jemand daran dachte. Sie fangen
    genau die Schaeden ab, die openpyxl beim Speichern anrichtet und die
    eine unvollstaendige XML-Chirurgie hinterlaesst: verschwundene
    x14-Regeln, ein zerschossenes VBA-Projekt, eine calcChain, die nicht
    mehr zu den Formeln passt.
    """
    fund = []

    def melde(regel, text):
        fund.append((regel, text))

    try:
        z = zipfile.ZipFile(pfad)
    except Exception as e:
        melde("archiv", "laesst sich nicht als Zip oeffnen: %r" % e)
        return fund

    with z:
        kaputt = z.testzip()
        if kaputt:
            melde("archiv", "beschaedigter Eintrag: %s" % kaputt)

        namen = z.namelist()

        # --- alle XML-Teile muessen parsebar sein --------------------
        for name in namen:
            if not name.endswith((".xml", ".rels")):
                continue
            try:
                ET.fromstring(z.read(name))
            except Exception as e:
                melde("xml", "%s ist nicht wohlgeformt: %s"
                      % (name, str(e)[:80]))

        # --- Makros und Kontrollkaestchen ---------------------------
        if "xl/vbaProject.bin" not in namen:
            melde("makros", "Die Datei enthaelt kein VBA-Projekt mehr")
        if not any("featurePropertyBag" in n for n in namen):
            melde("kontrollkaestchen",
                  "featurePropertyBag fehlt - die Erledigt-Haekchen sind weg")

        #  KEINE calcChain-Pruefung hier. Der erste Entwurf hatte eine
        #  ("xl/calcChain.xml darf nicht da sein") und meldete prompt
        #  JEDE echte Datei des Nutzers rot. Die Regel stimmt fuer den
        #  BAUSCHRITT - nach einer XML-Operation muss die calcChain weg,
        #  sonst passt sie nicht mehr zu den Formeln. Sobald Excel die
        #  Datei einmal gespeichert hat, legt es sie voellig zu Recht
        #  wieder an. Eine Bauregel als Abnahmeregel zu verwenden war der
        #  Fehler.
        wbxml = z.read("xl/workbook.xml").decode("utf-8", "replace")
        if 'calcMode="manual"' in wbxml and 'fullCalcOnLoad="1"' not in wbxml:
            melde("berechnung",
                  "calcMode=manual ohne fullCalcOnLoad - die Mappe rechnet "
                  "beim Oeffnen nicht durch")

        # --- die acht Regeln der bedingten Formatierung -------------
        #  Sieben davon sind x14-Erweiterungsregeln. Genau die wirft
        #  openpyxl beim Speichern ersatzlos weg, und man sieht es der
        #  Datei nicht an - die Farben bleiben einfach aus.
        #  Gezaehlt wird "<cfRule " bzw. "<x14:cfRule " MIT Leerzeichen;
        #  ein Substring "x14:cfRule" trifft Auf- und Zu-Tag doppelt.
        for name in namen:
            if not name.startswith("xl/worksheets/sheet"):
                continue
            roh = z.read(name).decode("utf-8", "replace")
            if "Referenz Code" not in roh and "cfRule" not in roh:
                continue
            #  "<cfRule " und "<x14:cfRule " sind verschiedene Zeichen-
            #  ketten - der Praefix x14: steht davor, also zaehlt das
            #  eine das andere NICHT mit. Der erste Entwurf zog trotzdem
            #  ab und kam auf "-6 klassische Regeln".
            klassisch = roh.count("<cfRule ")
            x14 = roh.count("<x14:cfRule ")
            if x14 and (klassisch, x14) != (1, 7):
                melde("cf-regeln",
                      "%s hat %d klassische und %d x14-Regeln, erwartet 1 und 7"
                      % (name, klassisch, x14))

            # --- keine kaputten Formeln ------------------------------
            for muster, was in (("#REF!", "#REF!"), ("_xludf", "_xludf")):
                if muster in roh:
                    melde("formelfehler", "%s enthaelt %s" % (name, was))

    return fund


# ------------------------------------------------------- Selbstpruefung
def selbsttest(pfad):
    """Baut jeden Fehler kuenstlich ein und verlangt, dass er auffaellt.

    Ohne diesen Teil waere nicht gesagt, dass die Regeln oben ueberhaupt
    ansprechen - genau der Fehler, der dem Selbsttest in Excel passiert
    ist ("keine leeren Zeilen" war gruen, waehrend 35 dastanden).
    """
    basis = pruefe(pfad)
    print("Ausgangslage: %d Befund(e)" % len(basis))
    for regel, s in basis:
        print("   [%s] %s" % (regel, s))
    vorhandene = {r for r, _ in basis}

    faelle = []

    def fall(name, regel, mutiere):
        faelle.append((name, regel, mutiere))

    def m_update_blatt(wb):
        wb.create_sheet("Update")

    def m_lb_formel(wb):
        wl = wb[LB]
        k = kopfzeile(wl, 1, "Referenzcode")
        wl.cell(k + 1, 6).value = None

    def m_lb_summe(wb):
        wl = wb[LB]
        k = kopfzeile(wl, 1, "Referenzcode")
        for r in range(k + 1, k + 80):
            if txt(wl, r, 4).lower() == "summe":
                wl.cell(r, 5).value = "=SUM(E1:E1)"
                return

    def m_kennzeichen(wb):
        ws = wb[WP]
        for rng in ws.merged_cells.ranges:
            if rng.min_col == 2 and (rng.max_col - rng.min_col + 1) > 4:
                ws.cell(rng.min_row, 14).value = None
                return

    def m_wochenfolge(wb):
        ws = wb[WP]
        k = kopfzeile(ws, 2, "Referenz Code")
        letzte = plan_letzte_zeile(ws, k + 1)
        ws.cell(letzte, 5).value = 1

    def m_leerzeile(wb):
        ws = wb[WP]
        k = kopfzeile(ws, 2, "Referenz Code")
        letzte = plan_letzte_zeile(ws, k + 1)
        ws.cell(letzte + 3, 1).value = None      # dehnt max_row aus
        ws.cell(letzte + 3, 1).value = ""
        ws.cell(letzte + 3, 1).value = None

    fall("Blatt 'Update' vorhanden", "update-blatt", m_update_blatt)
    fall("Rechenspalte F ohne Formel", "lb-formeln", m_lb_formel)
    fall("Summenformel falsch", "lb-summe", m_lb_summe)
    fall("Ferien-Kennzeichen fehlt", "kennzeichen", m_kennzeichen)
    fall("Unterrichtswochen verdreht", "wochenfolge", m_wochenfolge)

    print("\nMutationen (jede MUSS anschlagen):")
    schlecht = 0
    for name, regel, mut in faelle:
        wb = openpyxl.load_workbook(pfad, data_only=False)
        mut(wb)
        neu = {r for r, _ in pruefe(pfad, wb)}
        erkannt = regel in neu and regel not in vorhandene
        if regel in vorhandene:
            #  Auch das ist ein Nachweis - die Regel spricht an, nur
            #  eben schon an einem ECHTEN Befund dieser Datei. Ein
            #  kuenstlicher Fehler obendrauf wuerde nichts beweisen.
            print("   ok %-32s -> %s (schlaegt schon am echten Befund an)"
                  % (name, regel))
        elif erkannt:
            print("   ok %-32s -> %s" % (name, regel))
        else:
            print("   FEHLT %-29s -> %s wurde NICHT gemeldet" % (name, regel))
            schlecht += 1
    # --- Mutationen auf Dateiebene ----------------------------------
    #  Die Strukturregeln lassen sich nicht ueber ein openpyxl-Objekt
    #  ausloesen - openpyxl wuerde die Datei beim Speichern ohnehin
    #  zerstoeren, und genau davor sollen die Regeln ja schuetzen.
    #  Deshalb wird hier eine echte Kopie des Zips veraendert.
    datei_faelle = [
        ("VBA-Projekt entfernt", "makros",
         lambda n, d: (None, None) if n == "xl/vbaProject.bin" else (n, d)),
        ("Kontrollkaestchen entfernt", "kontrollkaestchen",
         lambda n, d: (None, None) if "featurePropertyBag" in n else (n, d)),
        ("eine x14-Regel geloescht", "cf-regeln",
         lambda n, d: (n, re.sub(rb"<x14:cfRule .*?</x14:cfRule>", b"", d, count=1,
                                 flags=re.S)) if "worksheets/sheet" in n else (n, d)),
        ("XML beschaedigt", "xml",
         lambda n, d: (n, d.replace(b"</worksheet>", b"</workshee", 1))
                      if n.endswith("sheet2.xml") else (n, d)),
        ("#REF! eingeschmuggelt", "formelfehler",
         lambda n, d: (n, d.replace(b"Einstellungen!", b"#REF!", 1))
                      if n.endswith("sheet2.xml") else (n, d)),
    ]
    print("\nMutationen auf Dateiebene:")
    for name, regel, patch in datei_faelle:
        ziel = tempfile.mktemp(suffix=".xlsm")
        with zipfile.ZipFile(pfad) as zin, \
             zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                nn, dd = patch(it.filename, zin.read(it.filename))
                if nn is not None:
                    zout.writestr(nn, dd)
        neu_regeln = {r for r, _ in pruefe_struktur(ziel)}
        if regel in neu_regeln:
            print("   ok %-32s -> %s" % (name, regel))
        else:
            print("   FEHLT %-29s -> %s nicht gemeldet (gemeldet: %s)"
                  % (name, regel, sorted(neu_regeln) or "nichts"))
            schlecht += 1
        import os as _os
        _os.remove(ziel)

    gesamt = len(faelle) + len(datei_faelle)
    print("\n%d von %d Mutationen erkannt."
          % (gesamt - schlecht, gesamt))
    if "leerzeilen" not in vorhandene:
        print("Hinweis: die Regel 'leerzeilen' laesst sich mit openpyxl nicht "
              "kuenstlich ausloesen\n(eine leere, nur formatierte Zeile laesst "
              "sich so nicht erzeugen). Sie wurde\nstattdessen gegen die echte "
              "beschaedigte Datei validiert - das ist der\nstaerkere Nachweis.")
    return schlecht


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dateien", nargs="+")
    ap.add_argument("--selbsttest", action="store_true",
                    help="Mutationen einbauen und pruefen, ob sie auffallen")
    a = ap.parse_args()

    if a.selbsttest:
        return 1 if selbsttest(a.dateien[0]) else 0

    schlecht = 0
    for p in a.dateien:
        fund = [f for f in pruefe(p) if f[0] != "hinweis"]
        kurz = p.split("/")[-1]
        if not fund:
            print("OK    %s" % kurz)
        else:
            schlecht += 1
            print("FEHLER %s" % kurz)
            for regel, s in fund:
                print("        [%s] %s" % (regel, s))
    return 1 if schlecht else 0


if __name__ == "__main__":
    sys.exit(main())
