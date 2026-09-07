#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Streift einen Authenticode-Signaturblock ab (Git-clean-Filter).

WARUM ES DIESE DATEI GIBT
    Makros_verteilen.ps1 wird lokal mit einem selbst erzeugten
    Testzertifikat signiert, damit es unter der eigenen
    Ausfuehrungsrichtlinie startet. Diese Signatur gehoert nicht ins
    Repository: sie gilt nur auf dem Rechner, der sie erzeugt hat, sie
    waere fuer jeden anderen wertlos, und sie erzeugt bei jedem Signieren
    eine Aenderung an einer Datei, die sich inhaltlich nicht geaendert hat.

    Sie von Hand vor jedem Commit zu entfernen waere die schlechteste
    Loesung - daran denkt niemand dauerhaft. Deshalb ein Git-Filter: die
    Arbeitsdatei bleibt signiert, im Repository landet sie ohne Block.

EINRICHTEN (einmalig je Arbeitsplatz)
    git config filter.signatur.clean "python Makros/signatur_entfernen.py"

    Die Zuordnung steht in .gitattributes. Ist der Filter nicht
    eingerichtet, reicht Git den Inhalt unveraendert durch - ein Klon
    ohne diese Einstellung funktioniert also weiterhin, nur dass ein
    dort signiertes Skript wieder mit Block eingecheckt wuerde.

AUFRUF
    Von Git, mit dem Dateiinhalt auf stdin und dem Ergebnis auf stdout.
"""
import re
import sys

#  PowerShell schreibt den Block ans Dateiende, jede Zeile als Kommentar,
#  und setzt eine LEERZEILE davor. Die muss mit weg: sonst unterscheidet
#  sich die Datei im Repository um genau einen Umbruch von der
#  unsignierten Fassung, und jedes Signieren erzeugte doch wieder eine
#  Scheinaenderung.
BLOCK = re.compile(
    rb"\r?\n# SIG # Begin signature block\r?\n"
    rb".*?# SIG # End signature block\r?\n?",
    re.S)


def main():
    roh = sys.stdin.buffer.read()
    sys.stdout.buffer.write(BLOCK.sub(b"", roh))
    return 0


if __name__ == "__main__":
    sys.exit(main())
