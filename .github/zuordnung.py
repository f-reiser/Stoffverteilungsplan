#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ordnet offenen Vorgaengen ihre bereits vorhandenen Branches zu.

Liest zwei Dateien, die der Workflow-Schritt vorher erzeugt hat:
    nummern.txt   je Zeile die Nummer eines Vorgangs mit Auftragslabel
    branches.txt  je Zeile der Name eines Branches im Repository

Schreibt nach GITHUB_OUTPUT:
    offen          Zahl der Vorgaenge mit Auftragslabel
    weiter_nr      Vorgang, an dem weitergearbeitet werden MUSS (sonst leer)
    weiter_branch  dessen Branch (sonst leer)

WARUM DAS HIER STEHT UND NICHT IM PROMPT
    Ein abgebrochener Lauf hinterlaesst einen Branch, waehrend das Issue
    sein Label behaelt. Wer den fortsetzt, darf nicht davon abhaengen, dass
    das Modell (a) dasselbe Issue nochmal auswaehlt und (b) beim Suchen
    dasselbe Namensschema raet, das der vorige Lauf beim Anlegen benutzt
    hat. Beides entscheidet jetzt der Runner, bevor das Modell startet -
    das kostet nichts und kann nicht anders ausgehen.
"""
import io
import os
import re
import sys

#  Das Schema aus git-branch-strategie. Streng geprueft, weil ein
#  abweichender Name genau den Fall unauffindbar macht, fuer den es
#  diese Datei gibt.
SCHEMA = re.compile(r"^issue-(\d+)-[A-Za-z0-9._-]+$")


def lies(pfad):
    try:
        with io.open(pfad, encoding="utf-8") as f:
            return [z.strip() for z in f if z.strip()]
    except OSError:
        return []


def zuordnen(nummern, branches):
    """(weiter_nr, weiter_branch, meldungen)"""
    offen = sorted({int(n) for n in nummern if n.isdigit()})
    je_nummer = {}
    fremd = []
    for b in branches:
        if b == "main":
            continue
        m = SCHEMA.match(b)
        if m:
            je_nummer.setdefault(int(m.group(1)), []).append(b)
        else:
            fremd.append(b)

    meldungen = []
    for b in fremd:
        #  Kein Abbruch: Ein Branch darf auch aus anderem Anlass entstehen.
        #  Sichtbar muss es trotzdem sein, denn wenn ein Lauf das Schema
        #  verlaesst, findet der naechste seine Arbeit nicht wieder.
        meldungen.append("Branch ausserhalb des Schemas issue-<nr>-<slug>: " + b)

    for nr, bs in sorted(je_nummer.items()):
        if nr not in offen:
            meldungen.append("Branch ohne offenen Auftrag (Label entfernt?): "
                             + ", ".join(sorted(bs)))

    kandidaten = [(nr, sorted(bs)) for nr, bs in sorted(je_nummer.items())
                  if nr in offen]
    if not kandidaten:
        return "", "", meldungen

    nr, bs = kandidaten[0]
    if len(bs) > 1:
        #  Zwei Branches zu einem Vorgang kann der Runner nicht aufloesen;
        #  raten waere schlimmer als abgeben.
        meldungen.append("Mehrere Branches zu Vorgang %d (%s) - keine "
                         "Vorgabe, das Modell entscheidet." % (nr, ", ".join(bs)))
        return "", "", meldungen
    if len(kandidaten) > 1:
        meldungen.append("Weitere angefangene Vorgaenge warten: "
                         + ", ".join(str(k[0]) for k in kandidaten[1:]))
    return str(nr), bs[0], meldungen


def main():
    nummern = lies("nummern.txt")
    branches = lies("branches.txt")
    nr, branch, meldungen = zuordnen(nummern, branches)
    offen = len({n for n in nummern if n.isdigit()})

    zeilen = ["Vorgaenge mit Auftragslabel: %d" % offen]
    if nr:
        zeilen.append("Fortsetzen: #%s auf %s" % (nr, branch))
    else:
        zeilen.append("Fortsetzen: nichts angefangen")
    zeilen += meldungen

    text = chr(10).join(zeilen)
    print(text)
    z = os.environ.get("GITHUB_STEP_SUMMARY")
    if z:
        io.open(z, "a", encoding="utf-8").write(text + chr(10))
    a = os.environ.get("GITHUB_OUTPUT")
    if a:
        io.open(a, "a", encoding="utf-8").write(
            "offen=%d%sweiter_nr=%s%sweiter_branch=%s%s"
            % (offen, chr(10), nr, chr(10), branch, chr(10)))
    return 0


def selbsttest():
    fehler = []

    def pruefe(name, fn, erwartet):
        try:
            ist = fn()
        except Exception as ex:
            ist = "AUSNAHME %s" % ex
        if ist != erwartet:
            fehler.append("%s: %r statt %r" % (name, ist, erwartet))

    #  Der Normalfall, um den es geht
    pruefe("angefangener Vorgang wird gesetzt",
           lambda: zuordnen(["7", "13"], ["main", "issue-7-leer-definition"])[:2],
           ("7", "issue-7-leer-definition"))

    #  Aehnliche Nummer darf nicht treffen
    pruefe("17 ist nicht 7",
           lambda: zuordnen(["7"], ["main", "issue-17-etwas"])[:2], ("", ""))

    #  Nichts angefangen
    pruefe("ohne Branch keine Vorgabe",
           lambda: zuordnen(["7", "13"], ["main"])[:2], ("", ""))

    #  Branch ohne Auftrag: keine Vorgabe, aber sichtbar
    pruefe("Branch ohne offenen Auftrag",
           lambda: zuordnen(["13"], ["main", "issue-9-alt"])[:2], ("", ""))
    pruefe("und er wird gemeldet",
           lambda: any("ohne offenen Auftrag" in m
                       for m in zuordnen(["13"], ["main", "issue-9-alt"])[2]), True)

    #  Schemaverstoss: nicht zugeordnet, aber gemeldet
    pruefe("fremdes Schema nicht zugeordnet",
           lambda: zuordnen(["7"], ["main", "fix-7-irgendwas"])[:2], ("", ""))
    pruefe("und gemeldet",
           lambda: any("ausserhalb des Schemas" in m
                       for m in zuordnen(["7"], ["main", "fix-7-irgendwas"])[2]), True)

    #  Zwei Branches zu einem Vorgang: lieber abgeben als raten
    pruefe("mehrdeutig gibt ab",
           lambda: zuordnen(["7"], ["issue-7-a", "issue-7-b"])[:2], ("", ""))

    #  Der kleinste angefangene Vorgang gewinnt, die anderen werden genannt
    pruefe("kleinste Nummer zuerst",
           lambda: zuordnen(["7", "13"], ["issue-13-x", "issue-7-y"])[:2],
           ("7", "issue-7-y"))
    pruefe("weitere werden genannt",
           lambda: any("Weitere angefangene" in m
                       for m in zuordnen(["7", "13"], ["issue-13-x", "issue-7-y"])[2]),
           True)

    #  Leere Eingaben duerfen nicht abstuerzen
    pruefe("gar nichts", lambda: zuordnen([], [])[:2], ("", ""))

    for f in fehler:
        print("FEHLER: " + f)
    print("%d von 11 Pruefungen bestanden." % (11 - len(fehler)))
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(selbsttest() if "--selbsttest" in sys.argv else main())
