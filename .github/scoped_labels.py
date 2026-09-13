#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Scoped Labels nach dem Vorbild von GitLab: "Scope::Wert"."""
import sys

TRENNER = "::"


def scope(label):
    """Der Teil vor "::", oder None, wenn das Label nicht gescoped ist."""
    kopf, trenner, _ = label.partition(TRENNER)
    if not trenner or not kopf:
        return None
    return kopf


def wert(label):
    """Der Teil hinter "::", oder None, wenn das Label nicht gescoped ist."""
    kopf, trenner, rest = label.partition(TRENNER)
    if not trenner or not kopf:
        return None
    return rest


def selbsttest():
    fehler = []

    def pruefe(label, erwarteter_scope, erwarteter_wert):
        s = scope(label)
        w = wert(label)
        if (s, w) != (erwarteter_scope, erwarteter_wert):
            fehler.append("%r: (%r, %r) statt (%r, %r)"
                          % (label, s, w, erwarteter_scope, erwarteter_wert))

    pruefe("Modell::Opus", "Modell", "Opus")
    pruefe("v::5", "v", "5")
    pruefe("Aufwand::extra hoch", "Aufwand", "extra hoch")
    #  Ein Scope, den weder geschuetzt.py noch modellwahl.py kennen - die
    #  Erkennung selbst darf trotzdem nichts von den beiden wissen.
    pruefe("Prioritaet::hoch", "Prioritaet", "hoch")

    #  Nicht gescoped: kein Trenner, oder Trenner am Rand.
    pruefe("Einarbeiten", None, None)
    pruefe("", None, None)
    pruefe("::ohne-scope", None, None)
    pruefe("ohne-wert::", "ohne-wert", "")

    #  Mehrfaches "::" gehoert komplett zum Wert - nur die erste Trennung zaehlt.
    pruefe("v::5::6", "v", "5::6")

    gesamt = 9
    for f in fehler:
        print("FEHLER: " + f)
    print("%d von %d Pruefungen bestanden." % (gesamt - len(fehler), gesamt))
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(selbsttest())
