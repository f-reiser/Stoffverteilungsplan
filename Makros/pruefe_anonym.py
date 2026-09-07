#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sucht produktive Inhalte in den Mappen unter Vorlage/ (Pruefebene A).

WARUM ES DIESE DATEI GIBT
    Vorlage/ ist der einzige Ort im Repository, an dem .xlsm liegen. Bis
    06.09.2026 war Vorlage/Referenzmappe.xlsm eine byteweise Kopie des
    echten Mathematik-Gym-10-Plans - mit Themen, Terminen, Notizen, dem
    Klarnamen der Lehrkraft, dem Schulnamen und dem vollstaendigen
    OneDrive-Pfad im versteckten Namen wpPdfOrdner. Aufgefallen ist das
    erst beim Vergleich der Pruefsummen, kurz vor dem ersten Push.

    Eine Zusage, "beim naechsten Mal daran zu denken", faengt das nicht
    ab. Eine Pruefung schon - und anders als ein Vorsatz laeuft sie bei
    jedem Push mit. Was einmal in der Git-Historie steht, bekommt man
    nicht mehr sauber heraus.

    Absichtlich NICHT gemeldet werden die Lernbereichsnamen und die
    Lehrplan-Codes (M10.5 &co.): die stehen so im LehrplanPLUS und sind
    oeffentlich. Gemeldet wird, was den Nutzer, seine Schule oder seine
    konkrete Unterrichtsplanung erkennbar macht.

AUFRUF
    python3 pruefe_anonym.py ../Vorlage/*.xlsm
    python3 pruefe_anonym.py --selbsttest ../Vorlage/*.xlsm

    Neue Fassung bauen: siehe anonymisiere.py
"""
import argparse
import glob
import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

import ci_ausgabe

#  (Name, Suchmuster, Beispiel fuer den Mutationstest)
#
#  Jede Regel hat ein Beispiel, und der Selbsttest baut genau dieses
#  Beispiel in eine saubere Mappe ein. Ohne diesen Nachweis waere von
#  keiner der Regeln gesagt, dass sie ueberhaupt anschlagen kann - genau
#  der Fehler, den dieses Projekt schon dreimal hatte.
#
#  Hier stehen nur die ALLGEMEINEN Spuren. Was den Betreiber oder seine
#  Schule benennt, gehoert nicht in ein oeffentliches Repository und
#  steht in MUSTER_LOKAL (siehe unten) - sonst veroeffentlicht
#  ausgerechnet die Pruefung das, was sie schuetzen soll.
SPUREN = [
    ("Seitenverweise ins Schulbuch", r"LS:\s*Seite\s*\d+", "LS: Seite 12 - 15"),
    ("Schulaufgaben-Planung", r"\d\.\s*(?:Schulaufgabe|kasL)", "1. Schulaufgabe"),
    ("eigene Kompetenzformulierung",
     r"Baumdiagramm|Monte-Carlo|Pfadregeln", "Pfadregeln"),
    ("eigene Notiz", r"Zweifelsfall ans Jahresende",
     "Wird im Zweifelsfall ans Jahresende geschoben"),
    ("Windows-Benutzerpfad", r"C:\\Users\\[A-Za-z0-9._-]+", r"C:\Users\beispiel"),
    ("OneDrive-Pfad", r"OneDrive", "OneDrive - Irgendeine Schule"),
    ("E-Mail-Adresse", r"[\w.+-]+@[\w-]+\.[\w.]+", "vorname.name@schule.de"),
]

#  Standortabhaengige Muster: Klarnamen, Schulnamen, alles Uebrige, was
#  nur dieser Betreiber kennt. Die Datei ist in .gitignore; daneben liegt
#  anonym_muster.local.json.vorlage als Beispiel.
MUSTER_LOKAL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "anonym_muster.local.json")


def _lokale_muster():
    """Liest die standortabhaengigen Muster; leere Liste, wenn keine da."""
    if not os.path.exists(MUSTER_LOKAL):
        return []
    with io.open(MUSTER_LOKAL, encoding="utf-8") as f:
        return [(e["name"], e["muster"], e["beispiel"]) for e in json.load(f)]


def alle_spuren():
    """Allgemeine und standortabhaengige Muster zusammen."""
    return SPUREN + _lokale_muster()


def pruefe(pfad):
    """Gibt [(Regel, Anzahl, Beispiel)] zurueck; leer = sauber."""
    z = zipfile.ZipFile(pfad)
    #  vbaProject.bin ist eine Binaerdatei und wird ausgelassen: dort
    #  steht der Quelltext der Module, und der gehoert ohnehin als .bas
    #  ins Repository.
    text = "\n".join(z.read(n).decode("utf-8", "replace")
                     for n in z.namelist()
                     if (n.endswith(".xml") or n.endswith(".rels"))
                     and "vbaProject" not in n)
    z.close()
    fund = []
    for name, muster, _ in alle_spuren():
        treffer = re.findall(muster, text, re.I)
        if treffer:
            fund.append((name, len(treffer), str(treffer[0])[:45]))
    return fund


def _mit_spur(quelle, ziel, text):
    """Kopie von `quelle`, in deren sharedStrings `text` zusaetzlich steht."""
    z = zipfile.ZipFile(quelle)
    sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    eintrag = "<si><t>%s</t></si>" % text.replace("&", "&amp;").replace("<", "&lt;")
    neu = sx.replace("</sst>", eintrag + "</sst>", 1)
    #  count und uniqueCount mitziehen, sonst ist die Mappe unstimmig -
    #  auch eine Mutation soll eine benutzbare Datei hinterlassen.
    for attr in ("count", "uniqueCount"):
        neu = re.sub(r'(<sst[^>]*?\b%s=")(\d+)(")' % attr,
                     lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                     neu, count=1)
    reihe = (["[Content_Types].xml"]
             + [n for n in z.namelist() if n != "[Content_Types].xml"])
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            daten = (neu.encode("utf-8") if n == "xl/sharedStrings.xml"
                     else z.read(n))
            out.writestr(z.getinfo(n), daten)
    z.close()


def selbsttest(pfade):
    sauber = [p for p in pfade if not pruefe(p)]
    if not sauber:
        print("Keine saubere Mappe als Ausgangsbasis - Selbsttest waere "
              "wertlos.\nZuerst anonymisiere.py laufen lassen.")
        return 1
    quelle = sauber[0]
    print("Ausgangsbasis: %s (sauber)\n" % os.path.basename(quelle))

    spuren = alle_spuren()
    if not _lokale_muster():
        print("Hinweis: keine standortabhaengigen Muster geladen (%s fehlt).\n"
              "         Klarnamen und Schulnamen werden damit NICHT geprueft."
              % os.path.basename(MUSTER_LOKAL))
    arbeit = tempfile.mkdtemp(prefix="pruefe_anonym_")
    ziel = os.path.join(arbeit, "mutiert.xlsm")
    offen = 0
    print("Mutationen (jede MUSS anschlagen):")
    for name, _, beispiel in spuren:
        _mit_spur(quelle, ziel, beispiel)
        gemeldet = {n for n, _, _ in pruefe(ziel)}
        if name in gemeldet:
            print("   ok    %s" % name)
        else:
            print("   FEHLT %-32s -> nicht gemeldet (gemeldet: %s)"
                  % (name, sorted(gemeldet) or "nichts"))
            offen += 1
    shutil.rmtree(arbeit, ignore_errors=True)
    print("\n%d von %d Mutationen erkannt." % (len(spuren) - offen, len(spuren)))
    return 1 if offen else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dateien", nargs="+")
    ap.add_argument("--selbsttest", action="store_true",
                    help="jede Spur einbauen und pruefen, ob sie auffaellt")
    a = ap.parse_args()

    pfade = []
    for muster in a.dateien:
        pfade += sorted(glob.glob(muster)) or [muster]
    pfade = [p for p in pfade if not os.path.basename(p).startswith("~$")]

    if a.selbsttest:
        return selbsttest(pfade)

    schlecht = 0
    for p in pfade:
        fund = pruefe(p)
        if not fund:
            print("SAUBER %s" % os.path.basename(p))
        else:
            schlecht += 1
            print("FUNDE  %s" % os.path.basename(p))
            for name, n, bsp in fund:
                zeile = "%s: %d x, z.B. %r" % (name, n, bsp)
                print("        %s" % zeile)
                ci_ausgabe.annotiere(zeile, datei=p)
    if schlecht:
        print("\nProduktive Inhalte gehoeren nicht ins Repository.\n"
              "Neue Fassung bauen: python3 anonymisiere.py <quelle> <ziel>")
    return 1 if schlecht else 0


if __name__ == "__main__":
    sys.exit(main())
