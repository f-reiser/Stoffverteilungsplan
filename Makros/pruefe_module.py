#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Auslieferungspruefung der VBA-Quelldateien (Pruefebene 7).

Prueft jede .bas und .txt im Makros-Ordner auf die Eigenschaften, die beim
Import in den VBA-Editor stimmen MUESSEN. Bis 05.09.2026 lief das als
Handgriff nach jeder Auslieferung - also gar nicht, sobald es jemand
vergisst.

WAS HIER SCHON SCHIEFGING
    - Das Edit-Werkzeug hat modAnleitung.bas als UTF-8 zurueckgeschrieben
      und dabei jeden Umlaut durch U+FFFD ersetzt. Gerettet wurde die Datei
      nur, weil die Vorgaengerfassung noch beim Nutzer lag.
    - Der Weg SendUserFile -> device_commit_files hat mindestens einmal eine
      cp1252-Datei als UTF-8 abgelegt ("ue" lag als C3 BC auf der Platte).
    - Ein Python-Heredoc hat "\\\\" in den VBA-Quelltext geschrieben, also
      zwei echte Backslashes statt einem. Zweimal passiert.

AUFRUF
    python3 pruefe_module.py mod*.bas *.txt
    python3 pruefe_module.py --selbsttest mod*.bas
"""
import sys, re, argparse, glob, os, shutil, tempfile
import ci_ausgabe

# Zeichen, die in cp1252 nicht existieren duerfen
UTF8_RESTE = [b"\xc3\xa4", b"\xc3\xb6", b"\xc3\xbc", b"\xc3\x84", b"\xc3\x96",
              b"\xc3\x9c", b"\xc3\x9f", b"\xc3\xa9", b"\xe2\x80"]

# Doppelt kodiert: cp1252-gelesener Text, in dem UTF-8-Bytes als Zeichen stehen
MOJIBAKE = ["Ã¤", "Ã¶", "Ã¼", "Ã„", "Ã–", "Ãœ", "ÃŸ", "â€"]


def pruefe(pfad):
    """Liefert eine Liste von (regel, text)-Befunden. Leer = in Ordnung."""
    fund = []

    def melde(regel, text):
        fund.append((regel, text))

    roh = open(pfad, "rb").read()

    # --- 1. Muss sich als cp1252 lesen lassen ------------------------
    try:
        t = roh.decode("cp1252")
    except UnicodeDecodeError as e:
        melde("kodierung", "laesst sich nicht als cp1252 lesen: %s" % e)
        return fund

    # --- 2. Keine UTF-8-Reste ----------------------------------------
    #  Eine als UTF-8 gespeicherte Datei laesst sich zwar als cp1252
    #  lesen, aber die Umlaute stehen dann als zwei Zeichen da.
    treffer = [s for s in UTF8_RESTE if s in roh]
    if treffer:
        melde("utf8", "enthaelt UTF-8-Bytefolgen: %s"
              % " ".join(s.hex() for s in treffer))

    # --- 3. Kein Ersatzzeichen ---------------------------------------
    #  U+FFFD heisst: hier stand mal ein Umlaut, und er ist weg.
    #
    #  ACHTUNG, hier lag ein Fehler in der Vorgaengerfassung dieser
    #  Pruefung, gefunden vom Mutationstest weiter unten: geprueft wurde
    #  auf das ZEICHEN U+FFFD im cp1252-gelesenen Text. Das kann es dort
    #  gar nicht geben - U+FFFD steht in der Datei als die drei Bytes
    #  EF BF BD, und die liest cp1252 als "ï¿½", also drei harmlose
    #  Zeichen. Die Pruefung konnte nie anschlagen.
    #  Richtig ist die Suche nach der BYTEFOLGE.
    if b"\xef\xbf\xbd" in roh:
        n = roh.count(b"\xef\xbf\xbd")
        zeile = roh[:roh.index(b"\xef\xbf\xbd")].count(b"\n") + 1
        melde("zerstoert", "%d x U+FFFD (Ersatzzeichen), erstes in Zeile %d - "
                           "hier stand ein Umlaut" % (n, zeile))

    # --- 4. Keine Doppelkodierung ------------------------------------
    for m in MOJIBAKE:
        if m in t:
            zeile = t[:t.index(m)].count("\n") + 1
            melde("doppelt", "doppelt kodiert (%r in Zeile %d)" % (m, zeile))
            break

    # --- 5. Keine doppelten Backslashes in Zeichenketten -------------
    #  VBA kennt kein Escaping: "\\" sind zwei echte Backslashes.
    for i, l in enumerate(t.split("\n"), 1):
        for m in re.finditer(r'"[^"\n]*\\\\[^"\n]*"', l):
            melde("backslash", "Zeile %d: doppelter Backslash in %s"
                  % (i, m.group(0)[:60]))

    # --- 6. Einheitliche Zeilenenden ---------------------------------
    #  Die ausgelieferten Module sind LF. Eine Mischung deutet auf eine
    #  Bearbeitung mit dem falschen Werkzeug hin.
    if b"\r\n" in roh and roh.count(b"\n") != roh.count(b"\r\n"):
        melde("zeilenenden", "Mischung aus CRLF und LF")

    # --- 7. Kopfzeile eines VBA-Moduls -------------------------------
    if pfad.endswith(".bas"):
        erste = t.split("\n", 1)[0].strip()
        modul = os.path.basename(pfad)[:-4]
        if not erste.startswith("Attribute VB_Name"):
            melde("kopf", "erste Zeile ist kein 'Attribute VB_Name': %r"
                  % erste[:50])
        elif ('"%s"' % modul) not in erste:
            melde("kopf", "Modulname in der Kopfzeile passt nicht zum "
                          "Dateinamen: %r" % erste)
        if not re.search(r"^Option Explicit\s*$", t, re.M):
            melde("kopf", "kein 'Option Explicit'")

    # --- 8. Datei endet mit Zeilenumbruch ----------------------------
    if roh and not roh.endswith(b"\n"):
        melde("abschluss", "letzte Zeile ohne Zeilenumbruch")

    return fund


# ----------------------------------------------------------- Selbsttest
def umlaut_zerstoeren(b):
    """Ersten Umlaut durch die UTF-8-Bytefolge von U+FFFD ersetzen."""
    for byte in (0xE4, 0xF6, 0xFC, 0xC4, 0xD6, 0xDC, 0xDF):
        if bytes([byte]) in b:
            return b.replace(bytes([byte]), b"\xef\xbf\xbd", 1)
    raise ValueError("kein Umlaut in der Ausgangsdatei")


def selbsttest(pfade):
    """Baut jeden Fehler kuenstlich ein und verlangt, dass er auffaellt.

    Ohne diesen Teil waere nicht gesagt, dass die Regeln oben ueberhaupt
    ansprechen - genau der Fehler, der dem Excel-Selbsttest passiert ist
    ("keine leeren Zeilen" war gruen, waehrend 35 dastanden).
    """
    quelle = None
    for p in pfade:
        if p.endswith(".bas") and not pruefe(p):
            quelle = p
            break
    if quelle is None:
        print("Kein fehlerfreies .bas als Ausgangsbasis gefunden.")
        return 1
    print("Ausgangsbasis: %s (fehlerfrei)\n" % quelle)
    roh = open(quelle, "rb").read()

    faelle = [
        ("als UTF-8 gespeichert", "utf8",
         lambda b: b.decode("cp1252").encode("utf-8")),
        #  Genau so sah die Datei aus, die das Edit-Werkzeug
        #  hinterlassen hat: ein Umlaut durch die UTF-8-Bytefolge des
        #  Ersatzzeichens ersetzt, der Rest unveraendert.
        ("Umlaut zerstoert (U+FFFD)", "zerstoert", umlaut_zerstoeren),
        ("doppelter Backslash", "backslash",
         lambda b: b.replace(b'Option Explicit',
                             b'Option Explicit\nPrivate Const X = "a\\\\b"', 1)),
        ("Zeilenenden gemischt", "zeilenenden",
         lambda b: b.replace(b"\n", b"\r\n", 3)),
        ("Kopfzeile fehlt", "kopf",
         lambda b: b.split(b"\n", 1)[1]),
        ("Option Explicit fehlt", "kopf",
         lambda b: b.replace(b"Option Explicit", b"' Option Explicit", 1)),
        ("Abschluss ohne Umbruch", "abschluss",
         lambda b: b.rstrip(b"\n")),
    ]

    #  ZWEI Fehler steckten in der Vorgaengerzeile
    #      tmp = "/tmp/_pruefe_module_test.bas"
    #
    #  1. "/tmp" gibt es nur auf Unix. In der Cowork-Linux-VM lief das,
    #     unter Windows brach der Selbsttest mit FileNotFoundError ab -
    #     und in der CI (Ubuntu) waere er nie aufgefallen. Eine Pruefung,
    #     die nur auf einem der beiden Wege laeuft, unterlaeuft den Zweck
    #     von pruefe_alles.py.
    #
    #  2. Der Fantasiename hat Regel 7 dauerhaft ausgeloest: sie
    #     vergleicht "Attribute VB_Name" mit dem DATEINAMEN, und
    #     "_pruefe_module_test" passt zu keinem Modul. Schon die
    #     unveraenderte Datei meldete damit "kopf" - also genau die
    #     Regel, die zwei der Mutationen unten erwarten. Beide waren
    #     gruen, ohne je etwas nachgewiesen zu haben.
    #     Deshalb behaelt die Kopie den Namen ihrer Quelldatei.
    arbeit = tempfile.mkdtemp(prefix="pruefe_module_")
    tmp = os.path.join(arbeit, os.path.basename(quelle))
    schlecht = 0
    print("Mutationen (jede MUSS anschlagen):")
    for name, regel, mutiere in faelle:
        try:
            kaputt = mutiere(roh)
        except Exception as e:
            print("   FEHLT %-28s -> Mutation nicht baubar: %r" % (name, e))
            schlecht += 1
            continue
        open(tmp, "wb").write(kaputt)
        regeln = {r for r, _ in pruefe(tmp)}
        if regel in regeln:
            print("   ok    %-28s -> %s" % (name, regel))
        else:
            print("   FEHLT %-28s -> %s wurde NICHT gemeldet (gemeldet: %s)"
                  % (name, regel, sorted(regeln) or "nichts"))
            schlecht += 1
    shutil.rmtree(arbeit, ignore_errors=True)
    print("\n%d von %d Mutationen erkannt." % (len(faelle) - schlecht, len(faelle)))
    return schlecht


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dateien", nargs="+")
    ap.add_argument("--selbsttest", action="store_true",
                    help="Mutationen einbauen und pruefen, ob sie auffallen")
    a = ap.parse_args()

    pfade = []
    for muster in a.dateien:
        pfade += sorted(glob.glob(muster)) or [muster]

    if a.selbsttest:
        return 1 if selbsttest(pfade) else 0

    schlecht = 0
    for p in pfade:
        fund = pruefe(p)
        if not fund:
            print("OK     %s" % os.path.basename(p))
        else:
            schlecht += 1
            print("FEHLER %s" % os.path.basename(p))
            for regel, s in fund:
                print("        [%s] %s" % (regel, s))
                #  Ausserhalb der CI wirkungslos. Traegt die Meldung eine
                #  Zeilennummer, wird daraus eine Annotation an genau
                #  dieser Zeile; sonst eine an der Datei.
                ci_ausgabe.annotiere(
                    "[%s] %s" % (p, s) if s.startswith("Zeile ") else s,
                    datei=p)
    return 1 if schlecht else 0


if __name__ == "__main__":
    sys.exit(main())
