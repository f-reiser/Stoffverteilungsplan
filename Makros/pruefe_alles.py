#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fuehrt alle automatisierbaren Pruefebenen aus - lokal wie in der CI.

WARUM ES DIESE DATEI GIBT
    Die Ebenen 1, 2, 3, 6, 7 und 8 waren bis zum 05.09.2026 zum Teil
    Gewohnheiten: Skripte, die waehrend einer Sitzung entstanden, ihren
    Zweck erfuellten und danach mit dem Sitzungsverzeichnis verschwanden.
    Wer die Arbeit spaeter uebernimmt, kann eine Gewohnheit nicht erben.
    Hier stehen sie als Code, mit einem einzigen Aufruf.

    Und weil die CI genau dieses Skript aufruft, koennen lokaler Lauf und
    Pipeline nicht auseinanderlaufen. Sobald jemand hier eine Pruefung
    ergaenzt, laeuft sie bei jedem Push mit.

AUFRUF
    python3 pruefe_alles.py                  # aus dem Makros-Ordner
    python3 pruefe_alles.py --mappen ../*.xlsm
    python3 pruefe_alles.py --ohne-mappen    # nur Quelltextebenen

RUECKGABE
    0 = alles gruen. Alles andere = mindestens eine Ebene rot.
"""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

import ci_ausgabe

HIER = os.path.dirname(os.path.abspath(__file__))

#  Archivdateien des Nutzers. Sie sind absichtlich auf altem Stand.
ARCHIV = re.compile(r"\s-\s?(v\d+|old|Test)\.xlsm$", re.I)


def lauf(titel, argv, cwd=None):
    """Ein Teilschritt. Gibt (name, rueckgabe) zurueck und druckt alles mit."""
    with ci_ausgabe.gruppe(titel):
        print("\n" + "=" * 70)
        print("  " + titel)
        print("=" * 70)
        print("$ " + " ".join(os.path.basename(a) if a.startswith("/") else a
                              for a in argv))
        #  Der Unterprozess schreibt direkt auf den Dateideskriptor,
        #  unsere prints gehen ueber Pythons Puffer. Ohne das Leeren
        #  landet die Kopfzeile NACH der Ausgabe des Schrittes - in der
        #  CI sogar ausserhalb der Klappgruppe.
        sys.stdout.flush()
        e = subprocess.run(argv, cwd=cwd or HIER)
    return (titel, e.returncode)


def erste_gefuellte(mappen):
    """Erste Mappe mit echten Plandaten, sonst None.

    WARUM DAS HIER STEHT
        Die Mutationstests der Ebenen 3 und 6 bauen Fehler in eine Mappe
        ein und verlangen, dass sie auffallen. In einer LEEREN Vorlage
        gibt es nichts zu verbiegen: beide Pruefer erkennen die leere
        Mappe und ueberspringen die Inhaltsregeln - die Mutation greift
        ins Nichts und der Test faellt durch.

        Beim ersten Lauf gegen die echte Vorlage war genau das der Fall
        (6 von 10 bzw. 0 von 4). Der Test hat sich richtig verhalten;
        falsch war die Referenzmappe. Damit dieser Unterschied nicht
        wieder als raetselhaftes Rot ankommt, wird er hier benannt.
    """
    try:
        import warnings
        warnings.filterwarnings("ignore")
        import openpyxl
    except ImportError:
        return mappen[0] if mappen else None

    for p in mappen:
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=False)
        except Exception:
            continue
        try:
            if "Wochenplan" not in wb.sheetnames:
                continue
            ws = wb["Wochenplan"]
            gefuellt = 0
            for r, zeile in enumerate(ws.iter_rows(min_row=4, max_row=60,
                                                   min_col=7, max_col=7,
                                                   values_only=True), 4):
                if zeile[0] not in (None, ""):
                    gefuellt += 1
            if gefuellt >= 5:
                return p
        finally:
            wb.close()
    return None


def nur_versionierte(pfade):
    """Filtert heraus, was .gitignore ausschliesst.

    Ebene A soll beantworten, ob etwas Produktives INS REPOSITORY
    geraet. Eigene Sicherungen, die im selben Ordner liegen
    ("... old.xlsm"), enthalten selbstverstaendlich echte Daten - sie
    mitzupruefen erzeugt ein Rot, das mit dem Repository nichts zu tun
    hat, und eine dauerhaft rote Pruefung liest bald niemand mehr.

    Ohne Git faellt die Filterung weg; dann wird alles geprueft.
    """
    wurzel = os.path.dirname(HIER)
    #  Git erwartet Pfade relativ zur Wurzel mit Schraegstrichen. Mit
    #  "..\\Vorlage\\x.xlsm" findet check-ignore nichts und meldet das
    #  auch nicht als Fehler - es filtert dann stillschweigend nichts.
    rel = {p: os.path.relpath(os.path.abspath(p), wurzel).replace("\\", "/")
           for p in pfade}
    #  Bewusst BYTES statt text=True: im Textmodus haengt Python unter
    #  Windows an jede Zeile ein \r. Git sucht dann nach "…xlsm\r",
    #  findet die Ausnahmeregel der Whitelist nicht mehr und meldet auch
    #  die echte Referenzmappe als ignoriert - der Filter haette
    #  ausgerechnet die Datei uebersprungen, um die es geht.
    try:
        e = subprocess.run(["git", "check-ignore", "-z", "--stdin"], cwd=wurzel,
                           input=b"\0".join(p.encode("utf-8")
                                            for p in rel.values()),
                           capture_output=True)
    except OSError:
        return pfade
    ignoriert = {z.decode("utf-8") for z in e.stdout.split(b"\0") if z}
    behalten = [p for p in pfade if rel[p] not in ignoriert]
    for p in pfade:
        if p not in behalten:
            print("Hinweis: %s ist in .gitignore - nicht Teil des "
                  "Repositorys, wird uebersprungen." % os.path.basename(p))
    return behalten


def module_bereitstellen():
    """Legt die Module in ein Arbeitsverzeichnis, notfalls mit Konfig-Vorlage.

    `modKonfig.bas` enthaelt das Blattschutz-Kennwort und gehoert deshalb
    NICHT ins Repository. Ohne dieses Modul kennt aber kein anderes die
    Konstante SCHUTZ_PW, und `vbacheck.py` meldet an 14 Stellen zu Recht
    "nicht deklariert". Deshalb liegt daneben `modKonfig.bas.vorlage` mit
    demselben Namen und einem Platzhalter-Kennwort - fuer die Pruefung
    reicht das, ausgeliefert wird sie nie.

    Rueckgabe: (Verzeichnis, Liste der .bas, Aufraeumfunktion)
    """
    echte = sorted(glob.glob(os.path.join(HIER, "mod*.bas")))
    namen = {os.path.basename(p) for p in echte}
    if "modKonfig.bas" in namen:
        return HIER, echte, lambda: None

    vorlage = os.path.join(HIER, "modKonfig.bas.vorlage")
    if not os.path.exists(vorlage):
        print("WARNUNG: weder modKonfig.bas noch modKonfig.bas.vorlage "
              "gefunden - SCHUTZ_PW ist nirgends deklariert.")
        return HIER, echte, lambda: None

    tmp = tempfile.mkdtemp(prefix="pruefe_alles_")
    for p in echte:
        shutil.copy2(p, tmp)
    shutil.copy2(vorlage, os.path.join(tmp, "modKonfig.bas"))
    print("Hinweis: modKonfig.bas fehlt (enthaelt das Kennwort und gehoert "
          "nicht ins Repository).\n"
          "         Fuer die Pruefung wird modKonfig.bas.vorlage eingesetzt.")
    return tmp, sorted(glob.glob(os.path.join(tmp, "mod*.bas"))), \
        lambda: shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mappen", nargs="*", default=None,
                    help="zu pruefende .xlsm; ohne Angabe nur Vorlage/")
    ap.add_argument("--alle-mappen", action="store_true",
                    help="zusaetzlich alle aktuellen .xlsm im Projektordner "
                         "(Archive mit ' - v1', ' -v2', ' - old', ' - Test' "
                         "bleiben aussen vor)")
    ap.add_argument("--ohne-mappen", action="store_true",
                    help="Ebenen 2, 3, 6 und 8 ueberspringen")
    a = ap.parse_args()

    verz, module, aufraeumen = module_bereitstellen()
    textdateien = sorted(glob.glob(os.path.join(HIER, "*_Modul.txt")))

    if not module:
        print("Keine mod*.bas gefunden - falscher Ordner?")
        return 2

    ergebnisse = []
    try:
        #  Zuerst die Ausgabeschicht selbst. Wenn die Meldewege kaputt
        #  sind, ist jede Meldung der folgenden Ebenen fragwuerdig.
        ergebnisse.append(lauf(
            "Werkzeug  Ausgabeschicht fuer die CI (ci_ausgabe.py)",
            [sys.executable, os.path.join(HIER, "ci_ausgabe.py"),
             "--selbsttest"]))

        # ---- Quelltextebenen: laufen immer, brauchen keine Mappe -----
        ergebnisse.append(lauf(
            "Ebene 1  statische Pruefung der Module (vbacheck.py)",
            [sys.executable, os.path.join(HIER, "vbacheck.py")] + module))

        ergebnisse.append(lauf(
            "Ebene 1  Mutationstest ueber vbacheck selbst",
            [sys.executable, os.path.join(HIER, "vbacheck.py"),
             "--selbsttest"] + module))

        ergebnisse.append(lauf(
            "Ebene 7  Auslieferungspruefung der Quelldateien (Kodierung)",
            [sys.executable, os.path.join(HIER, "pruefe_module.py")]
            + module + textdateien))

        ergebnisse.append(lauf(
            "Ebene 7  Mutationstest ueber die Auslieferungspruefung",
            [sys.executable, os.path.join(HIER, "pruefe_module.py"),
             "--selbsttest"] + module))

        #  Anonymitaet der Repo-Mappen. Bewusst mit FESTER Dateiliste:
        #  geprueft wird ausschliesslich Vorlage/, denn nur das liegt im
        #  Repository. Die eigenen Plaene des Nutzers enthalten
        #  selbstverstaendlich echte Daten - sie hier mitzupruefen waere
        #  eine dauerhaft rote Meldung ohne Aussage.
        repo_mappen = nur_versionierte(
            sorted(glob.glob(os.path.join(HIER, "..", "Vorlage", "*.xlsm"))))
        if repo_mappen:
            ergebnisse.append(lauf(
                "Ebene A  keine produktiven Inhalte in Vorlage/",
                [sys.executable, os.path.join(HIER, "pruefe_anonym.py")]
                + repo_mappen))

            ergebnisse.append(lauf(
                "Ebene A  Mutationstest ueber die Anonymitaetspruefung",
                [sys.executable, os.path.join(HIER, "pruefe_anonym.py"),
                 "--selbsttest"] + repo_mappen))

        # ---- Mappenebenen -------------------------------------------
        if a.ohne_mappen:
            mappen = []
        elif a.mappen is not None:
            mappen = a.mappen
        else:
            #  Voreinstellung: NUR die Referenzmappe. Im Projektordner
            #  liegen daneben die Archive des Nutzers (" - v1", " -v2",
            #  " - old", " - Test"). Die sind absichtlich auf altem
            #  Stand und muessen rot sein - eine Pruefung, die dauerhaft
            #  rot ist, liest bald niemand mehr. Mit --alle-mappen
            #  kommen die aktuellen Plaene dazu, die Archive nie.
            mappen = sorted(glob.glob(os.path.join(HIER, "..", "Vorlage",
                                                   "*.xlsm")))
            if a.alle_mappen:
                weitere = sorted(glob.glob(os.path.join(HIER, "..", "*.xlsm")))
                mappen += [m for m in weitere if not ARCHIV.search(
                    os.path.basename(m))]
            mappen = [m for m in mappen
                      if not os.path.basename(m).startswith("~$")]

        if not mappen:
            print("\n" + "=" * 70)
            print("  Ebenen 2, 3, 6, 8 uebersprungen - keine .xlsm gefunden")
            print("=" * 70)
            print("Die fertigen Mappen sind Binaerdateien und gehoeren nicht\n"
                  "in die Historie. Damit die Dateiebenen trotzdem bei jedem\n"
                  "Push laufen, gehoert GENAU EINE Referenzmappe ins Repo:\n"
                  "    Vorlage/Stoffverteilungsplan_Template.xlsm\n"
                  "Sie wird nur ausgetauscht, wenn sich die Vorlage aendert.\n"
                  "Die eigenen Plaene pruefst du mit --alle-mappen.")
        else:
            ergebnisse.append(lauf(
                "Ebenen 2+8  Abnahme der fertigen Mappen (pruefe_datei.py)",
                [sys.executable, os.path.join(HIER, "pruefe_datei.py")]
                + mappen))

            ergebnisse.append(lauf(
                "Ebene 6  Formelkonsistenz (pruefe_formeln.py)",
                [sys.executable, os.path.join(HIER, "pruefe_formeln.py")]
                + mappen))

            #  Die Mutationstests bekommen GENAU EINE gefuellte Mappe -
            #  siehe erste_gefuellte().
            gefuellt = erste_gefuellte(mappen)
            if gefuellt is None:
                print("\n" + "=" * 70)
                print("  Ebene 3 und Ebene 6: Mutationstests nicht durchfuehrbar")
                print("=" * 70)
                print("Keine der geprueften Mappen enthaelt Plandaten. In einer\n"
                      "leeren Vorlage laesst sich kein Fehler einbauen, der\n"
                      "auffallen koennte - die Mutationstests wuerden reihenweise\n"
                      "durchfallen, ohne dass an den Pruefungen etwas falsch ist.\n"
                      "\n"
                      "Abhilfe: eine GEFUELLTE Mappe als Referenz nach Vorlage/\n"
                      "legen (eine eingefrorene Kopie eines echten Plans, die nur\n"
                      "bewusst ausgetauscht wird). Die leere Vorlage darf daneben\n"
                      "liegen bleiben - die Abnahme prueft sie mit.")
                ergebnisse.append(
                    ("Ebene 3+6  Mutationstests (keine gefuellte Referenzmappe)", 1))
            else:
                print("\nReferenzmappe fuer die Mutationstests: %s"
                      % os.path.basename(gefuellt))
                ergebnisse.append(lauf(
                    "Ebene 3  Mutationstest ueber die Abnahme",
                    [sys.executable, os.path.join(HIER, "pruefe_datei.py"),
                     "--selbsttest", gefuellt]))

                ergebnisse.append(lauf(
                    "Ebene 6  Mutationstest ueber die Formelkonsistenz",
                    [sys.executable, os.path.join(HIER, "pruefe_formeln.py"),
                     "--selbsttest", gefuellt]))
    finally:
        aufraeumen()

    # ---- Zusammenfassung --------------------------------------------
    print("\n" + "=" * 70)
    print("  ZUSAMMENFASSUNG")
    print("=" * 70)
    rot = 0
    for titel, code in ergebnisse:
        print("  %-6s %s" % ("ROT" if code else "gruen", titel))
        rot += 1 if code else 0
    print()
    if rot:
        print("%d von %d Schritten rot." % (rot, len(ergebnisse)))
    else:
        print("Alle %d Schritte gruen." % len(ergebnisse))
    print()
    print("Was hier NICHT geprueft werden kann:")
    print("  Ebene 4/5  Selbsttest und Mutationstest in echtem Excel -")
    print("             brauchen Excel; laufen ueber die Schaltflaechen")
    print("             'Selbsttest' und 'Selbsttest pruefen' in der Mappe.")
    print("  Ebene 9    Fremde Gegenlese - braucht ein Sprachmodell.")
    print("             Vorgehen: Doku/Ebene9_Gegenlese.md")

    #  Dieselbe Tabelle noch einmal fuer die Laufseite in GitHub. Ohne
    #  das muesste man das Log aufklappen, um zu sehen, WAS rot ist.
    #  Ausserhalb der CI passiert hier nichts.
    for titel, code in ergebnisse:
        if code:
            ci_ausgabe.annotiere("Prueflauf rot: %s" % titel)
    ci_ausgabe.zusammenfassung("\n".join(
        ["## Prueflauf", "", "| Ergebnis | Schritt |", "|---|---|"]
        + ["| %s | %s |" % (":x: rot" if code else ":white_check_mark: gruen",
                            titel) for titel, code in ergebnisse]
        + ["",
           "**%d von %d Schritten rot.**" % (rot, len(ergebnisse)) if rot
           else "**Alle %d Schritte gruen.**" % len(ergebnisse),
           "",
           "Nicht automatisierbar: Ebene 4/5 (brauchen echtes Excel), "
           "Ebene 9 (fremde Gegenlese, braucht ein Sprachmodell)."]))
    return 1 if rot else 0


if __name__ == "__main__":
    sys.exit(main())
