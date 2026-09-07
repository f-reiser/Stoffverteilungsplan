#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tauscht die eingebetteten Schullogos einer Mappe gegen eigene aus.

WARUM ES DIESE DATEI GIBT
    Die Logos stecken FEST in der Mappe (Formen wpKopfLogo_<n>), nicht in
    einem Ordner daneben - der frueher vorhandene Apparat, der sie aus
    "Bilder" suchte, ist bewusst entfallen (siehe LIESMICH.txt). Das ist
    im Betrieb richtig, macht aber jede neue Fassung der Vorlage zum
    Problem: sie bringt die Platzhalter mit, und die eigenen Logos waeren
    von Hand wieder einzubetten.

    Dieses Skript ersetzt die Bilddaten im entpackten Zip. Geometrie,
    Namen und Anker der Formen bleiben unberuehrt - nur die Bytes der
    Bilder werden getauscht. Damit ist das Zurueckbranden nach einem
    Update ein Aufruf statt Handarbeit in Excel.

AUFRUF
    python3 logos_einsetzen.py <mappe.xlsm> <logo1.png> [<logo2.png> ...]

    Die Logos werden der Reihe nach eingesetzt (image1, image2, ...), in
    derselben Reihenfolge wie die Schulliste in modKonfig.bas.

VORHER LESEN
    Das Seitenverhaeltnis sollte zum Original passen: der Selbsttest
    verlangt Breite/Hoehe > 3, die mitgelieferten Platzhalter sind
    714x80 bzw. 775x80. Ein quadratisches Logo laesst den Test rot werden
    und sieht im Titelblock gequetscht aus.
"""
import os
import struct
import sys
import zipfile


def masse(daten):
    """Breite und Hoehe eines PNG; None bei allem anderen."""
    if daten[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    return struct.unpack(">II", daten[16:24])


def einsetzen(mappe, bilder, ziel=None):
    z = zipfile.ZipFile(mappe)
    medien = sorted(n for n in z.namelist() if n.startswith("xl/media/"))
    if not medien:
        raise SystemExit("In %s stecken keine Bilder." % mappe)
    if len(bilder) > len(medien):
        raise SystemExit("%d Bilder uebergeben, aber nur %d in der Mappe (%s)."
                         % (len(bilder), len(medien),
                            ", ".join(os.path.basename(m) for m in medien)))
    if len(bilder) < len(medien):
        print("Hinweis: %d von %d Bildern werden ersetzt, der Rest bleibt."
              % (len(bilder), len(medien)))

    neu = {}
    for m, pfad in zip(medien, bilder):
        daten = open(pfad, "rb").read()
        alt, jetzt = masse(z.read(m)), masse(daten)
        if alt and jetzt:
            print("  %-16s %sx%s  ->  %-24s %sx%s"
                  % (os.path.basename(m), alt[0], alt[1],
                     os.path.basename(pfad), jetzt[0], jetzt[1]))
            if jetzt[0] / jetzt[1] <= 3:
                print("     WARNUNG: Verhaeltnis %.2f - der Selbsttest verlangt "
                      "mehr als 3." % (jetzt[0] / jetzt[1]))
        else:
            print("  %s -> %s" % (os.path.basename(m), os.path.basename(pfad)))
        neu[m] = daten

    ziel = ziel or mappe
    reihe = (["[Content_Types].xml"]
             + [n for n in z.namelist() if n != "[Content_Types].xml"])
    tmp = ziel + ".neu"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            out.writestr(z.getinfo(n), neu.get(n) or z.read(n))
    z.close()
    os.replace(tmp, ziel)
    print("geschrieben: %s" % ziel)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    einsetzen(sys.argv[1], sys.argv[2:])
