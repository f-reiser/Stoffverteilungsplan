#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ausgabeschicht fuer GitHub Actions - lokal unveraendert, in der CI reich.

WARUM ES DIESE DATEI GIBT
    GitHub wertet den Rueckgabewert aus, mehr nicht. Das reicht fuer
    "rot oder gruen", aber nicht fuer die Frage, die danach kommt:
    WO ist es rot? Wer das in einem 500-Zeilen-Log suchen muss, liest bald
    gar nicht mehr nach - und eine Pruefung, die niemand liest, ist keine.

    GitHub kennt dafuer drei Kanaele, die kein Skript von sich aus nutzt:

      Annotationen   ::error file=X,line=N::Text
                     erscheinen oben im Lauf UND an der Codezeile in der
                     Dateiansicht eines Pull Requests.
      Job Summary    Markdown in der Datei, auf die $GITHUB_STEP_SUMMARY
                     zeigt. Wird als Tabelle auf der Laufseite gerendert.
      Log-Gruppen    ::group:: / ::endgroup:: - klappt lange Ausgaben zu.

WAS SICH LOKAL AENDERT: NICHTS
    Ausserhalb der CI (kein GITHUB_ACTIONS in der Umgebung) geben alle
    Funktionen hier None zurueck und drucken nichts. Die Pruefskripte
    behalten jeden ihrer bisherigen print-Aufrufe; sie bekommen nur eine
    zusaetzliche Zeile. Damit kann der lokale Lauf nicht verschlechtert
    werden - und das war die Bedingung, unter der diese Datei entstand.

WARUM KEIN unittest / pytest
    Die Pruefskripte sind keine Unit-Tests, sondern Dateipruefer, die auch
    von Hand als Auslieferungs-Gate laufen (`vbacheck.py mod*.bas`). Ein
    Test-Runner haette diesen Aufruf zerschlagen und einen ZWEITEN
    Ausfuehrungsweg geschaffen - genau das Auseinanderlaufen von
    Arbeitsplatz und Pipeline, gegen das `pruefe_alles.py` gebaut ist.

AUFRUF
    python3 ci_ausgabe.py --selbsttest    # prueft sich selbst
"""
import argparse
import contextlib
import os
import re
import sys
import tempfile

HIER = os.path.dirname(os.path.abspath(__file__))
WURZEL = os.path.dirname(HIER)

#  "[pfad] Zeile 42: text" - so melden vbacheck.py und pruefe_module.py.
#  Der Pfad wird nicht-gierig gelesen, sonst frisst er bei einem
#  Windows-Pfad ("C:\...") den eigenen Doppelpunkt mit auf.
MELDUNG = re.compile(r"^\[(.+?)\]\s*Zeile\s+(\d+):\s*(.*)$")


def ist_ci():
    """True, wenn wir in GitHub Actions laufen."""
    return os.environ.get("GITHUB_ACTIONS") == "true"


def escape_daten(s):
    """Escaping fuer den TEXT einer Workflow-Anweisung.

    Das Prozentzeichen MUSS zuerst ersetzt werden - sonst trifft die
    Ersetzung das '%' der eigenen Ersatzsequenz noch einmal und aus
    '%0A' wird '%250A'.
    """
    return s.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")


def escape_eigenschaft(s):
    """Escaping fuer einen EIGENSCHAFTSWERT (file=..., line=...).

    Zusaetzlich ':' und ',' - beide trennen in der Anweisungssyntax.
    Ein Windows-Pfad beginnt mit 'C:'; unescaped verwirft GitHub die
    ganze Annotation stillschweigend.
    """
    return escape_daten(s).replace(":", "%3A").replace(",", "%2C")


def zerlege_meldung(zeile):
    """'[pfad] Zeile 42: text' -> (pfad, 42, text); sonst None."""
    m = MELDUNG.match(zeile.strip())
    if not m:
        return None
    return (m.group(1), int(m.group(2)), m.group(3))


def repo_relativ(pfad):
    """Absoluter Pfad -> Pfad relativ zur Projektwurzel, mit '/'.

    Die Rueckabbildung ueber den Dateinamen ist kein Luxus: sobald
    modKonfig.bas fehlt, kopiert `pruefe_alles.py` die Module in ein
    Temp-Verzeichnis - und weil das Kennwortmodul nie im Repository
    liegt, ist das in der CI IMMER der Fall. Ohne diesen Zweig zeigt
    dort keine einzige Annotation auf eine echte Datei.
    """
    ap = os.path.abspath(pfad)
    if os.path.normcase(ap).startswith(os.path.normcase(WURZEL) + os.sep):
        return os.path.relpath(ap, WURZEL).replace("\\", "/")
    name = os.path.basename(ap)
    if name and os.path.exists(os.path.join(HIER, name)):
        return "Makros/" + name
    return None


def fehlerzeile(text, datei=None, zeile=None, art="error"):
    """Baut die Workflow-Anweisung. Reine Funktion, druckt nicht."""
    teile = []
    if datei:
        teile.append("file=" + escape_eigenschaft(datei))
    if zeile:
        teile.append("line=" + escape_eigenschaft(str(zeile)))
    kopf = art + (" " + ",".join(teile) if teile else "")
    return "::%s::%s" % (kopf, escape_daten(text))


def annotiere(meldung, art="error", datei=None):
    """Macht aus einer Pruefmeldung eine Annotation - nur in der CI.

    Nimmt die Zeile, die das Skript ohnehin schon druckt. Gibt ausserhalb
    der CI None zurueck und druckt nichts.

    `datei` greift nur, wenn die Meldung selbst keinen Pfad mitbringt.
    Ebene 7 meldet so: ihre Kodierungsbefunde gelten der ganzen Datei und
    haben oft gar keine Zeilennummer - ohne diesen Weg haenge gerade die
    teuerste Fehlerklasse des Projekts (zerstoerte Umlaute) an keiner Datei.
    """
    if not ist_ci():
        return None
    zerlegt = zerlege_meldung(meldung)
    if zerlegt:
        pfad, nr, text = zerlegt
        rel = repo_relativ(pfad)
        z = fehlerzeile(text, rel, nr, art) if rel else fehlerzeile(text, art=art)
    else:
        z = fehlerzeile(meldung, repo_relativ(datei) if datei else None,
                        None, art)
    print(z)
    return z


@contextlib.contextmanager
def gruppe(titel):
    """Klappbare Log-Gruppe in der CI, sonst nur eine Ueberschrift."""
    if ist_ci():
        print("::group::" + titel)
    try:
        yield
    finally:
        if ist_ci():
            print("::endgroup::")


def zusammenfassung(markdown):
    """Haengt Markdown ans Job Summary. Ausserhalb der CI wirkungslos.

    Rueckgabe: True, wenn geschrieben wurde.
    """
    ziel = os.environ.get("GITHUB_STEP_SUMMARY")
    if not ziel:
        return False
    with open(ziel, "a", encoding="utf-8") as f:
        f.write(markdown.rstrip("\n") + "\n")
    return True


# ----------------------------------------------------------------------
#  Die Pruefungen. Jede gibt None zurueck (gut) oder einen Grund.
# ----------------------------------------------------------------------

@contextlib.contextmanager
def _als_ci(an=True):
    alt = os.environ.get("GITHUB_ACTIONS")
    os.environ["GITHUB_ACTIONS"] = "true" if an else "nein"
    try:
        yield
    finally:
        if alt is None:
            os.environ.pop("GITHUB_ACTIONS", None)
        else:
            os.environ["GITHUB_ACTIONS"] = alt


def _gleich(ist, soll):
    return None if ist == soll else "erwartet %r, bekam %r" % (soll, ist)


def _p_prozent_zuerst():
    return _gleich(escape_daten("50%\nx"), "50%25%0Ax")


def _p_umbrueche():
    return _gleich(escape_daten("a\r\nb"), "a%0D%0Ab")


def _p_eigenschaft_doppelpunkt():
    return _gleich(escape_eigenschaft("C:\\a,b"), "C%3A\\a%2Cb")


def _p_eigenschaft_erbt_daten():
    return _gleich(escape_eigenschaft("100%"), "100%25")


def _p_zerlegen_normal():
    return _gleich(zerlege_meldung("[modKopf.bas] Zeile 42: irgendwas ist faul"),
                   ("modKopf.bas", 42, "irgendwas ist faul"))


def _p_zerlegen_windowspfad():
    return _gleich(zerlege_meldung(r"[C:\a\modKopf.bas] Zeile 7: kaputt"),
                   (r"C:\a\modKopf.bas", 7, "kaputt"))


def _p_zerlegen_fremdes():
    return _gleich(zerlege_meldung("OK - keine Beanstandungen"), None)


def _p_relativ_im_projekt():
    return _gleich(repo_relativ(os.path.join(HIER, "vbacheck.py")),
                   "Makros/vbacheck.py")


def _p_relativ_aus_temp():
    return _gleich(repo_relativ(os.path.join(tempfile.gettempdir(),
                                             "vbacheck.py")),
                   "Makros/vbacheck.py")


def _p_fehlerzeile_form():
    return _gleich(fehlerzeile("kaputt", "Makros/vbacheck.py", 42),
                   "::error file=Makros/vbacheck.py,line=42::kaputt")


def _p_fehlerzeile_escapet_text():
    return _gleich(fehlerzeile("a\nb"), "::error::a%0Ab")


def _p_fehlerzeile_escapet_eigenschaft():
    return _gleich(fehlerzeile("x", "a,b.bas"), "::error file=a%2Cb.bas::x")


def _p_annotiere_lokal_still():
    with _als_ci(False):
        return _gleich(annotiere("[modKopf.bas] Zeile 3: kaputt"), None)


def _p_annotiere_in_ci():
    with _als_ci(True):
        ist = annotiere("[%s] Zeile 3: kaputt"
                        % os.path.join(tempfile.gettempdir(), "vbacheck.py"))
    return _gleich(ist, "::error file=Makros/vbacheck.py,line=3::kaputt")


def _p_annotiere_mit_datei():
    """Meldung ohne eigenen Pfad muss trotzdem an ihrer Datei haengen."""
    with _als_ci(True):
        ist = annotiere("enthaelt UTF-8-Bytefolgen",
                        datei=os.path.join(tempfile.gettempdir(),
                                           "vbacheck.py"))
    return _gleich(ist, "::error file=Makros/vbacheck.py::enthaelt "
                        "UTF-8-Bytefolgen")


def _p_zusammenfassung_ohne_ziel():
    alt = os.environ.pop("GITHUB_STEP_SUMMARY", None)
    try:
        return _gleich(zusammenfassung("# egal"), False)
    finally:
        if alt is not None:
            os.environ["GITHUB_STEP_SUMMARY"] = alt


PRUEFUNGEN = [
    ("Prozentzeichen wird zuerst ersetzt", _p_prozent_zuerst),
    ("Zeilenumbrueche werden ersetzt", _p_umbrueche),
    ("Eigenschaft: Doppelpunkt und Komma", _p_eigenschaft_doppelpunkt),
    ("Eigenschaft escapet auch Prozent", _p_eigenschaft_erbt_daten),
    ("Meldung zerlegen", _p_zerlegen_normal),
    ("Meldung zerlegen: Windows-Pfad", _p_zerlegen_windowspfad),
    ("Fremde Zeile ist kein Treffer", _p_zerlegen_fremdes),
    ("Pfad relativ zur Wurzel", _p_relativ_im_projekt),
    ("Pfad aus dem Temp-Verzeichnis", _p_relativ_aus_temp),
    ("Annotation hat die richtige Form", _p_fehlerzeile_form),
    ("Annotation escapet den Text", _p_fehlerzeile_escapet_text),
    ("Annotation escapet den Dateinamen", _p_fehlerzeile_escapet_eigenschaft),
    ("Ausserhalb der CI bleibt es still", _p_annotiere_lokal_still),
    ("In der CI entsteht eine Annotation", _p_annotiere_in_ci),
    ("Meldung ohne Pfad haengt an ihrer Datei", _p_annotiere_mit_datei),
    ("Job Summary ohne Ziel schreibt nicht", _p_zusammenfassung_ohne_ziel),
]


# ----------------------------------------------------------------------
#  Mutationen. Jede baut einen Fehler ein und MUSS auffallen.
#  Projektregel: jede Pruefung hat eine Mutation, oder sie zaehlt nicht.
#
#  Alle Mutationen hier sind plattformunabhaengig. Eine erste Fassung
#  mutierte die Backslash-Ersetzung in repo_relativ - die kann auf Linux
#  gar nicht anschlagen, weil es dort keine Backslashes in Pfaden gibt.
#  Der Mutationstest waere in der CI rot gewesen, ohne dass an der
#  Pruefung etwas falsch ist.
# ----------------------------------------------------------------------

def _mut_annotiere_ohne_datei(meldung, art="error", datei=None):
    """Wie annotiere, aber der datei-Zweig fehlt.

    Bewusst chirurgisch: eine Mutation, die annotiere komplett ersetzt,
    faellt schon der allgemeinen Annotationspruefung auf - dann waere
    unbewiesen, ob die Pruefung des datei-Zweigs ueberhaupt anschlaegt.
    """
    if not ist_ci():
        return None
    zerlegt = zerlege_meldung(meldung)
    if zerlegt:
        pfad, nr, text = zerlegt
        rel = repo_relativ(pfad)
        z = fehlerzeile(text, rel, nr, art) if rel else fehlerzeile(text, art=art)
    else:
        z = fehlerzeile(meldung, art=art)
    print(z)
    return z


MUTATIONEN = [
    ("Prozent zuletzt statt zuerst", "escape_daten",
     lambda s: s.replace("\r", "%0D").replace("\n", "%0A").replace("%", "%25")),
    ("Umbrueche gar nicht ersetzt", "escape_daten",
     lambda s: s.replace("%", "%25")),
    ("Eigenschaft ohne Doppelpunkt", "escape_eigenschaft",
     lambda s: escape_daten(s).replace(",", "%2C")),
    ("Eigenschaft ohne Prozent", "escape_eigenschaft",
     lambda s: s.replace(":", "%3A").replace(",", "%2C")),
    ("Zerlegen verliert die Zeilennummer", "zerlege_meldung",
     lambda z: (z.split("]")[0][1:], 0, z) if z.startswith("[") else None),
    ("Zerlegen nimmt jede Zeile an", "zerlege_meldung",
     lambda z: (z, 1, z)),
    ("Relativ gibt den absoluten Pfad", "repo_relativ",
     lambda p: os.path.abspath(p)),
    ("Relativ ohne Temp-Rueckabbildung", "repo_relativ",
     lambda p: (os.path.relpath(os.path.abspath(p), WURZEL).replace("\\", "/")
                if os.path.normcase(os.path.abspath(p)).startswith(
                    os.path.normcase(WURZEL) + os.sep) else None)),
    ("Annotation ohne Text-Escaping", "fehlerzeile",
     lambda text, datei=None, zeile=None, art="error": "::%s::%s" % (art, text)),
    ("Annotation ohne Eigenschafts-Escaping", "fehlerzeile",
     lambda text, datei=None, zeile=None, art="error":
         "::%s%s::%s" % (art, (" file=" + datei) if datei else "",
                         escape_daten(text))),
    ("annotiere druckt auch ausserhalb der CI", "annotiere",
     lambda meldung, art="error", datei=None: fehlerzeile(meldung, art=art)),
    ("annotiere ignoriert die uebergebene Datei", "annotiere",
     _mut_annotiere_ohne_datei),
    ("Job Summary schreibt ohne Ziel", "zusammenfassung",
     lambda markdown: True),
]


def _laufe_pruefungen():
    """Fuehrt alle Pruefungen aus - stumm.

    Die Ausgabe wird weggeworfen, weil `annotiere` im Erfolgsfall druckt.
    Ohne das schriebe der Selbsttest in der CI echte ::error-Anweisungen
    ins Log, und GitHub haengte einem gruenen Lauf erfundene Fehler an.
    Die Pruefungen geben ihr Urteil zurueck, nicht aus - Drucken ist hier
    also nichts, was verloren gehen koennte.
    """
    schlecht = []
    for name, fn in PRUEFUNGEN:
        try:
            with open(os.devnull, "w") as still:
                with contextlib.redirect_stdout(still):
                    grund = fn()
        except Exception as e:
            grund = "Ausnahme: %r" % (e,)
        if grund:
            schlecht.append((name, grund))
    return schlecht


def selbsttest():
    print("Pruefungen (muessen alle bestehen):")
    schlecht = _laufe_pruefungen()
    fehlerhaft = {n: g for n, g in schlecht}
    for name, _ in PRUEFUNGEN:
        if name in fehlerhaft:
            print("   FEHLER %-38s -> %s" % (name, fehlerhaft[name]))
        else:
            print("   ok     %s" % name)
    if schlecht:
        print("\nAusgangslage ist nicht sauber - Mutationstest waere wertlos.")
        return 1

    print("\nMutationen (jede MUSS auffallen):")
    hier = sys.modules[__name__]
    offen = 0
    for name, funktion, ersatz in MUTATIONEN:
        echt = getattr(hier, funktion)
        setattr(hier, funktion, ersatz)
        try:
            gefunden = _laufe_pruefungen()
        finally:
            setattr(hier, funktion, echt)
        if gefunden:
            print("   ok     %-42s -> %s" % (name, gefunden[0][0]))
        else:
            print("   FEHLT  %-42s -> keine Pruefung hat angeschlagen" % name)
            offen += 1

    print("\n%d von %d Mutationen erkannt."
          % (len(MUTATIONEN) - offen, len(MUTATIONEN)))
    return 1 if offen else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--selbsttest", action="store_true")
    a = ap.parse_args()
    if a.selbsttest:
        return selbsttest()
    print(__doc__)
    return 0


if __name__ == "__main__":
    sys.exit(main())
