#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ordnet offenen Vorgaengen ihre bereits vorhandenen Branches zu.

Liest zwei Dateien, die der Workflow-Schritt vorher erzeugt hat:
    nummern.txt   je Zeile die Nummer eines Vorgangs mit Auftragslabel
    branches.txt  je Zeile der Name eines Branches im Repository

Schreibt nach GITHUB_OUTPUT:
    offen       Zahl der Vorgaenge mit Auftragslabel
    fortsetzen  angefangene Vorgaenge als "nr auf branch", mit "; " getrennt
                (sonst leer). Hoechstens so viele, wie ein Durchgang
                ueberhaupt annehmen darf - der Rest wartet auf den naechsten.

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

#  Deckungsgleich mit "Hoechstens drei Vorgaenge pro Durchgang" aus
#  github-issue-workflow. Mehr anzureichen waere sinnlos: Der Lauf duerfte
#  sie ohnehin nicht annehmen.
HOECHSTENS = 3


def lies(pfad):
    try:
        with io.open(pfad, encoding="utf-8") as f:
            return [z.strip() for z in f if z.strip()]
    except OSError:
        return []


def zuordnen(nummern, branches):
    """(paare, meldungen) - paare als Liste (nr, branch), hoechstens HOECHSTENS"""
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

    paare = []
    for nr, bs in sorted(je_nummer.items()):
        if nr not in offen:
            continue
        if len(bs) > 1:
            #  Zwei Branches zu einem Vorgang kann der Runner nicht
            #  aufloesen; raten waere schlimmer als abgeben.
            meldungen.append("Mehrere Branches zu Vorgang %d (%s) - keine "
                             "Vorgabe, das Modell entscheidet."
                             % (nr, ", ".join(sorted(bs))))
            continue
        paare.append((nr, bs[0]))

    if len(paare) > HOECHSTENS:
        meldungen.append("Angefangen sind %d Vorgaenge; %d passen in einen "
                         "Durchgang, der Rest wartet: %s"
                         % (len(paare), HOECHSTENS,
                            ", ".join(str(n) for n, _ in paare[HOECHSTENS:])))
        paare = paare[:HOECHSTENS]
    return paare, meldungen


def main():
    nummern = lies("nummern.txt")
    branches = lies("branches.txt")
    paare, meldungen = zuordnen(nummern, branches)
    offen = len({n for n in nummern if n.isdigit()})
    fortsetzen = "; ".join("%d auf %s" % p for p in paare)

    zeilen = ["Vorgaenge mit Auftragslabel: %d" % offen]
    zeilen.append("Fortsetzen: " + (fortsetzen or "nichts angefangen"))
    zeilen += meldungen

    text = chr(10).join(zeilen)
    print(text)
    z = os.environ.get("GITHUB_STEP_SUMMARY")
    if z:
        io.open(z, "a", encoding="utf-8").write(text + chr(10))
    a = os.environ.get("GITHUB_OUTPUT")
    if a:
        io.open(a, "a", encoding="utf-8").write(
            "offen=%d%sfortsetzen=%s%s" % (offen, chr(10), fortsetzen, chr(10)))
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

    def paare(nummern, branches):
        return zuordnen(nummern, branches)[0]

    def meldet(nummern, branches, teil):
        return any(teil in m for m in zuordnen(nummern, branches)[1])

    #  Der Normalfall, um den es geht
    pruefe("angefangener Vorgang wird gereicht",
           lambda: paare(["7", "13"], ["main", "issue-7-leer-definition"]),
           [(7, "issue-7-leer-definition")])

    #  Mehrere angefangene: alle, aufsteigend
    pruefe("mehrere, aufsteigend",
           lambda: paare(["7", "13"], ["issue-13-x", "issue-7-y"]),
           [(7, "issue-7-y"), (13, "issue-13-x")])

    #  Mehr als in einen Durchgang passen: gekappt und gemeldet
    viele = [str(n) for n in (3, 5, 7, 9)]
    zweige = ["issue-%d-x" % n for n in (3, 5, 7, 9)]
    pruefe("auf HOECHSTENS gekappt",
           lambda: len(paare(viele, zweige)), HOECHSTENS)
    pruefe("und der Rest genannt",
           lambda: meldet(viele, zweige, "der Rest wartet"), True)

    #  Aehnliche Nummer darf nicht treffen
    pruefe("17 ist nicht 7", lambda: paare(["7"], ["main", "issue-17-etwas"]), [])

    #  Nichts angefangen
    pruefe("ohne Branch nichts", lambda: paare(["7", "13"], ["main"]), [])

    #  Branch ohne Auftrag: nicht reichen, aber sichtbar
    pruefe("Branch ohne offenen Auftrag",
           lambda: paare(["13"], ["main", "issue-9-alt"]), [])
    pruefe("und er wird gemeldet",
           lambda: meldet(["13"], ["main", "issue-9-alt"], "ohne offenen Auftrag"), True)

    #  Schemaverstoss: nicht zugeordnet, aber gemeldet
    pruefe("fremdes Schema nicht zugeordnet",
           lambda: paare(["7"], ["main", "fix-7-irgendwas"]), [])
    pruefe("und gemeldet",
           lambda: meldet(["7"], ["main", "fix-7-irgendwas"], "ausserhalb des Schemas"),
           True)

    #  Zwei Branches zu einem Vorgang: dieser faellt raus, andere bleiben
    pruefe("mehrdeutiger faellt raus, anderer bleibt",
           lambda: paare(["7", "13"], ["issue-7-a", "issue-7-b", "issue-13-x"]),
           [(13, "issue-13-x")])
    pruefe("Mehrdeutigkeit wird gemeldet",
           lambda: meldet(["7"], ["issue-7-a", "issue-7-b"], "Mehrere Branches"), True)

    #  Leere Eingaben duerfen nicht abstuerzen
    pruefe("gar nichts", lambda: paare([], []), [])

    for f in fehler:
        print("FEHLER: " + f)
    print("%d von 13 Pruefungen bestanden." % (13 - len(fehler)))
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(selbsttest() if "--selbsttest" in sys.argv else main())
