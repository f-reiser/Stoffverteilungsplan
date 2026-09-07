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
    Das Seitenverhaeltnis muss zum Original passen - die mitgelieferten
    Platzhalter sind 714x80 und 775x80.

    Darauf ist zu achten, weil es NICHTS nachtraeglich prueft: Der
    Selbsttest misst mit LogoVerhaeltnis die Breite und Hoehe der FORM,
    nicht die des Bildes. Die Form bleibt hier unveraendert. Ein
    quadratisches Logo wird also in den flachen Rahmen gezerrt, sieht im
    Titelblock gequetscht aus - und der Selbsttest bleibt gruen. Dieses
    Skript warnt deshalb selbst, wenn das Verhaeltnis stark abweicht.
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
        #  Eine leere oder fremdformatige Datei stillschweigend
        #  durchzuwinken waere der schlimmste Ausgang: Der Aufruf meldete
        #  Erfolg, in der Mappe bliebe das alte Logo stehen - und genau
        #  darauf verlaesst sich jemand, der gerade entbrandet.
        if not daten:
            raise SystemExit("%s ist leer." % pfad)
        jetzt = masse(daten)
        if jetzt is None:
            raise SystemExit("%s ist kein PNG. Die Mappe deklariert die "
                             "Endung png; ein anderes Format kann Excel "
                             "beim Oeffnen beanstanden." % pfad)
        alt = masse(z.read(m))
        print("  %-16s %s  ->  %-24s %sx%s"
              % (os.path.basename(m),
                 "%sx%s" % alt if alt else "?",
                 os.path.basename(pfad), jetzt[0], jetzt[1]))
        v = jetzt[0] / jetzt[1]
        if v <= 3:
            print("     WARNUNG: Verhaeltnis %.2f. Die Form bleibt flach, das "
                  "Bild wird gezerrt - und kein Test merkt es." % v)
        neu[m] = daten

    ziel = ziel or mappe
    reihe = (["[Content_Types].xml"]
             + [n for n in z.namelist() if n != "[Content_Types].xml"])
    tmp = ziel + ".neu"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in reihe:
            #  Bewusst "n in neu" statt "neu.get(n) or ...": ein falsy
            #  Wert waere sonst still durch das Original ersetzt worden.
            out.writestr(z.getinfo(n), neu[n] if n in neu else z.read(n))
    z.close()
    os.replace(tmp, ziel)
    print("geschrieben: %s" % ziel)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    einsetzen(sys.argv[1], sys.argv[2:])
