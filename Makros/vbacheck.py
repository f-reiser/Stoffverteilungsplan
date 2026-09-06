#!/usr/bin/env python3
"""Statische Pruefung fuer die .bas-Module des Stoffverteilungsplans."""
import re, sys, collections
import ci_ausgabe

VBA_KEYWORDS = set("""
in shift before after type alertstyle operator formula1 formula2 output append
destination paste link output count password contents scenarios
allowformattingcells allowformattingcolumns allowformattingrows allowinsertingcolumns
allowinsertingrows allowinsertinghyperlinks allowdeletingcolumns allowdeletingrows
allowsorting allowfiltering allowusingpivottables structure windows drawingobjects
collection

if then else elseif end select case for each next do loop while wend with
dim redim set let const public private static sub function property get
exit on error resume goto gosub return true false nothing empty null and or
not xor mod is like to step byval byref optional paramarray as new call
option explicit base compare module preserve type enum declare lib alias
attribute me thisworkbook application selection activesheet activewindow
integer long single double string boolean date variant object currency byte
""".split())

# Explizit bekannte Excel-/Office-Konstanten. Bewusst KEINE Praefix-Regel:
# ein Tippfehler wie xlNoRestriction (statt xlNoRestrictions) wuerde sonst
# durchrutschen und erst zur Laufzeit als "Variable nicht definiert" auffallen.
OFFICE_CONSTANTS = set("""
xlcalculation xlcalculationmanual xlcalculationautomatic
xlup xldown xltoleft xltoright xlshiftup xlshiftdown
xlvalidatelist xlvalidalertstop xlbetween
xlcontinuous xlthin xlsolid xlnone
xlcenter xlleft xlright xltop xlbottom xljustify
xlfreefloating xlmoveandsize xlmove
xlpasteformats xlpastevalues xlpasteall
xlnorestrictions xlunlockedcells xlnoselection
xledgebottom xledgetop xledgeleft xledgeright xllandscape xlportrait
xllistseparator
xlpapera4 xlpapera3 xltypepdf xlqualitystandard xlqualityminimum
xlprintnocomments xlprinterrorsdisplayed
xlnormalview xlpagebreakpreview xlpagelayoutview
xlmovedown xlvalue xlsheetvisible xlsheethidden
xlmedium xlthick xlhairline xlinsidevertical xlinsidehorizontal
xldiagonaldown xldiagonalup xlautomatic
msotrue msofalse
msofiledialogfilepicker msofiledialogfolderpicker
msoshapecross msoshapeisoscelestriangle msoshapemathplus msoshapemathminus
msoshaperectangle msoshaperoundedrectangle
msobringtofront msosendtoback
msoaligncenter msoalignleft msoalignright
msoanchormiddle msoanchortop
msoautosizenone msoautosizeshapetofittext
""".split())


BUILTIN = set("""
msgbox vbcrlf vbcr vblf vbtab vbok vbcancel vbyes vbno vbquestion vbexclamation
vbinformation vbokcancel vbyesno vbdefaultbutton2 vbboolean vbempty
range cells rows columns worksheet worksheets workbook shape shapes
workbooks activeworkbook windows
format formats left right mid trim ucase lcase len instr instrrev val cstr clng
cdbl cbool cdate cint isnumeric isdate isempty isnull isarray isobject iserror
varttype vartype typename array ubound lbound err error date now time year month day
dateserial datediff dateadd weekday rgb iif choose switch abs int fix sgn round
freefile open close print input line write kill dir beep debug string space timer
rtrim ltrim replace split join strconv asc chr chrw ascw sqr exp log rnd
environ createobject getobject shell doevents fileexists filedialog
strcomp strreverse cvar cvdate clnglng partition
xlup xldown xltoleft xltoright xlcalculationmanual xlcalculationautomatic
xlvalidatelist xlvalidalertstop xlbetween xlcontinuous xlthin xlsolid
xlcenter xlleft xlright xlfreefloating xlmoveandsize xlpasteformats
xllistseparator xlmovedown
mso msotrue msofalse
""".split())


def strip_line(l):
    # Kommentare entfernen (Strings beachten)
    out = []
    in_str = False
    i = 0
    while i < len(l):
        ch = l[i]
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif ch == "'" and not in_str:
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(text):
    """Fortsetzungszeilen (_ am Ende) zusammenfassen; (nr, code) zurueck."""
    raw = text.splitlines()
    res = []
    buf = ""
    start = 0
    for n, l in enumerate(raw, 1):
        code = strip_line(l).rstrip()
        if not buf:
            start = n
        if code.endswith(" _"):
            buf += code[:-1]
            continue
        buf += code
        res.append((start, buf))
        buf = ""
    if buf:
        res.append((start, buf))
    return res


def check(path):
    text = open(path, encoding="cp1252").read()
    lines = logical_lines(text)
    problems = []

    # --- 1. Deklarationsteil ---------------------------------------
    first_proc = None
    for n, l in lines:
        if re.match(r"^\s*(public\s+|private\s+|friend\s+)?(static\s+)?(sub|function|property)\b", l, re.I):
            first_proc = n
            break
    if first_proc is None:
        pass          # reines Konstantenmodul (modKonfig) - in Ordnung
    else:
        for n, l in lines:
            if n >= first_proc:
                break
            s = l.strip()
            if not s:
                continue
            if not re.match(r"^(option|attribute|dim|public|private|const|type|enum|declare|end type|end enum|#)", s, re.I):
                problems.append(f"Zeile {n}: Anweisung im Deklarationsteil: {s!r}")

    # --- 1b. Deklaration NACH der ersten Prozedur -------------------
    #  Die Gegenrichtung zu 1, und die fehlte bis 05.09.2026: eine
    #  modulweite Deklaration zwischen zwei Prozeduren. VBA laesst das
    #  nicht zu und meldet beim Kompilieren "Nach End Sub, End Function
    #  oder End Property koennen nur Kommentare stehen" - gefolgt von
    #  "Variable nicht definiert" an jeder Verwendung.
    #  Passiert ist es, weil drei Konstanten dort standen, wo sie
    #  inhaltlich hingehoeren: direkt vor der Prozedur, die sie
    #  benutzt. Naheliegend, lesbar - und in VBA verboten.
    #  Innerhalb einer Prozedur ist Dim natuerlich erlaubt, deshalb
    #  zaehlt hier nur, was auf Modulebene steht.
    in_proc = False
    for n, l in lines:
        s = l.strip()
        if re.match(r"^(public\s+|private\s+|friend\s+)?(static\s+)?(sub|function|property)\b", s, re.I):
            in_proc = True
            continue
        if re.match(r"^end\s+(sub|function|property)\b", s, re.I):
            in_proc = False
            continue
        if in_proc or first_proc is None or n < first_proc:
            continue
        if re.match(r"^(dim|const|static)\b", s, re.I) or \
           re.match(r"^(public|private)\s+(?!sub\b|function\b|property\b|declare\b|type\b|enum\b)", s, re.I):
            problems.append(
                f"Zeile {n}: modulweite Deklaration nach der ersten Prozedur - "
                f"das laesst VBA nicht zu, alles Modulweite gehoert nach oben: {s!r}")

    # --- 2. Blockbalance -------------------------------------------
    depth = 0
    stack = []
    for n, l in lines:
        s = l.strip()
        low = s.lower()
        if re.match(r"^(public |private |friend |static )*(sub|function|property\s+(get|let|set))\b", low):
            stack.append(("proc", n)); continue
        if re.match(r"^end\s+(sub|function|property)\b", low):
            if not stack or stack[-1][0] != "proc":
                problems.append(f"Zeile {n}: End Sub/Function ohne Anfang")
            else:
                stack.pop()
            continue
        if re.match(r"^(if\b.*\bthen\s*$)", low):
            stack.append(("if", n)); continue
        if re.match(r"^select\s+case\b", low):
            stack.append(("select", n)); continue
        if re.match(r"^(for)\b", low) and not re.match(r"^for\b.*\bnext\b", low):
            stack.append(("for", n)); continue
        if re.match(r"^with\b", low):
            stack.append(("with", n)); continue
        if re.match(r"^do\b", low):
            stack.append(("do", n)); continue
        if re.match(r"^end\s+if\b", low):
            if not stack or stack[-1][0] != "if":
                problems.append(f"Zeile {n}: End If ohne If (offen: {stack[-1] if stack else None})")
            else: stack.pop()
            continue
        if re.match(r"^end\s+select\b", low):
            if not stack or stack[-1][0] != "select":
                problems.append(f"Zeile {n}: End Select ohne Select")
            else: stack.pop()
            continue
        if re.match(r"^end\s+with\b", low):
            if not stack or stack[-1][0] != "with":
                problems.append(f"Zeile {n}: End With ohne With (offen: {stack[-1] if stack else None})")
            else: stack.pop()
            continue
        if re.match(r"^next\b", low):
            if not stack or stack[-1][0] != "for":
                problems.append(f"Zeile {n}: Next ohne For (offen: {stack[-1] if stack else None})")
            else: stack.pop()
            continue
        if re.match(r"^loop\b", low) or re.match(r"^wend\b", low):
            if not stack or stack[-1][0] != "do":
                problems.append(f"Zeile {n}: Loop ohne Do")
            else: stack.pop()
            continue
    if stack:
        problems.append(f"Am Dateiende noch offen: {stack}")

    # --- 3. FormatConditions ---------------------------------------
    for n, l in lines:
        if "formatcondition" in l.lower():
            problems.append(f"Zeile {n}: Zugriff auf FormatConditions - verboten")


    # --- 5. Verbundene Zellen: Vollbereich quer durch B..M ----------
    # Ferienzeilen im Wochenplan sind ueber B:M verbunden. Ein Bereich,
    # der von der ersten bis zur letzten Planzeile laeuft und dabei
    # Spalten INNERHALB von B..M anspricht, schneidet diese Verbindungen
    # an -> Laufzeitfehler 1004. Erlaubt ist nur der volle Bereich B..M
    # (der jede Verbindung ganz enthaelt) oder blockweises Arbeiten.
    for n, l in lines:
        # Nur Lesen/Schneiden ist unproblematisch - gemeint sind
        # Zuweisungen an Eigenschaften des Bereichs.
        if "Intersect(" in l:
            continue
        for m in re.finditer(
                r'Range\(\s*"([A-Z]{1,2})"\s*&\s*WP_FIRST_ROW\s*&\s*":([A-Z]{1,2})"\s*&\s*(\w+)',
                l):
            c1, c2 = m.group(1), m.group(2)
            if len(c1) == 1 and len(c2) == 1 and 'B' <= c1 <= 'M' and 'B' <= c2 <= 'M':
                if not (c1 == 'B' and c2 == 'M'):
                    problems.append(
                        f"Zeile {n}: Bereich {c1}..{c2} ueber alle Planzeilen - "
                        f"schneidet die verbundenen Ferienzeilen (B:M) an. "
                        f"Blockweise arbeiten!")

    # --- 4. Deklarierte Namen sammeln ------------------------------
    declared = set()
    for n, l in lines:
        s = l.strip()
        m = re.match(r"^(?:public|private|friend|global)?\s*(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+([A-Za-z_]\w*)", s, re.I)
        if m:
            declared.add(m.group(1).lower())
            # Parameter
            args = s[s.find("(") + 1:s.rfind(")")] if "(" in s else ""
            for a in args.split(","):
                am = re.search(r"(?:byval\s+|byref\s+|optional\s+|paramarray\s+)*([A-Za-z_]\w*)", a.strip(), re.I)
                if am:
                    declared.add(am.group(1).lower())
            continue
        m = re.match(r"^(?:public|private|global|dim|static)\s+(?:const\s+)?(.+)$", s, re.I)
        if m and not re.match(r"^(public|private)\s+(sub|function|declare|type|enum|const\s+)?$", s, re.I):
            body = m.group(1)
            if re.match(r"^(sub|function|property|declare|type|enum)\b", body, re.I):
                continue
            for part in re.split(r",(?![^()]*\))", body):
                pm = re.match(r"\s*([A-Za-z_]\w*)", part)
                if pm:
                    declared.add(pm.group(1).lower())
            continue
        m = re.match(r"^const\s+(.+)$", s, re.I)
        if m:
            for part in m.group(1).split(","):
                pm = re.match(r"\s*([A-Za-z_]\w*)", part)
                if pm:
                    declared.add(pm.group(1).lower())
            continue
        m = re.match(r"^redim\s+(?:preserve\s+)?([A-Za-z_]\w*)", s, re.I)
        if m:
            declared.add(m.group(1).lower())

    return problems, declared, lines


MODULES = {"modwochenplan", "modkalender", "modsteuerung", "modanleitung",
           "modschutz", "modkopf", "modkonfig", "modselbsttest", "moduebernahme",
           "modstart"}

PROC_START = re.compile(
    r"^\s*(?:(public|private|friend)\s+)?(?:static\s+)?"
    r"(sub|function|property\s+(?:get|let|set))\s+([A-Za-z_]\w*)", re.I)
PROC_END = re.compile(r"^\s*end\s+(sub|function|property)\b", re.I)


def deklarierte_namen(zeile):
    """Namen, die diese eine Zeile deklariert (Dim/Const/Static/ReDim)."""
    s = zeile.strip()
    namen = []
    m = re.match(r"^(?:public|private|global|dim|static)\s+(?:const\s+)?(.+)$", s, re.I)
    if m:
        body = m.group(1)
        if re.match(r"^(sub|function|property|declare|type|enum)\b", body, re.I):
            return []
        for teil in re.split(r",(?![^()]*\))", body):
            pm = re.match(r"\s*([A-Za-z_]\w*)", teil)
            if pm:
                namen.append(pm.group(1).lower())
        return namen
    m = re.match(r"^const\s+(.+)$", s, re.I)
    if m:
        for teil in m.group(1).split(","):
            pm = re.match(r"\s*([A-Za-z_]\w*)", teil)
            if pm:
                namen.append(pm.group(1).lower())
        return namen
    m = re.match(r"^redim\s+(?:preserve\s+)?([A-Za-z_]\w*)", s, re.I)
    if m:
        namen.append(m.group(1).lower())
    return namen


def parameter_namen(kopfzeile):
    if "(" not in kopfzeile:
        return []
    args = kopfzeile[kopfzeile.find("(") + 1:kopfzeile.rfind(")")]
    namen = []
    for a in args.split(","):
        am = re.search(r"(?:byval\s+|byref\s+|optional\s+|paramarray\s+)*([A-Za-z_]\w*)",
                       a.strip(), re.I)
        if am:
            namen.append(am.group(1).lower())
    return namen


def zerlege(lines):
    """(modulweite Namen, oeffentliche Namen, [Prozeduren]) einer Datei.

    Eine Prozedur ist (name, startzeile, eigene Namen, [(nr, code)]).
    """
    modul_namen, oeffentlich, prozeduren = set(), set(), []
    aktuell = None
    for n, l in lines:
        m = PROC_START.match(l)
        if m and aktuell is None:
            sicht, art, nm = m.group(1), m.group(2), m.group(3)
            eigene = set(parameter_namen(l)) | {nm.lower()}
            aktuell = [nm, n, eigene, []]
            # Prozedurnamen sind modulweit; Private nur in diesem Modul.
            modul_namen.add(nm.lower())
            if (sicht or "public").lower() != "private":
                oeffentlich.add(nm.lower())
            continue
        if aktuell is not None:
            if PROC_END.match(l):
                prozeduren.append(tuple(aktuell))
                aktuell = None
                continue
            aktuell[2].update(deklarierte_namen(l))
            #  Sprungmarken ("Fail:", "Ende:") sind Namen wie jeder
            #  andere - "GoTo Fail" muss also erlaubt sein, "GoTo
            #  Nirgendwo" nicht. Vorher standen "fail" und "cleanfail"
            #  einfach als Woerter in VBA_KEYWORDS: damit war jede
            #  Sprungmarke mit einem anderen Namen ein Fehlalarm, und
            #  ein Sprung ins Leere fiel nicht auf. Genau das ist beim
            #  Label "Ende" passiert.
            lm = re.match(r"^([A-Za-z_]\w*):\s*(?:'.*)?$", l.strip())
            if lm:
                aktuell[2].add(lm.group(1).lower())
            aktuell[3].append((n, l))
            continue
        # Deklarationsteil
        s = l.strip()
        namen = deklarierte_namen(s)
        modul_namen.update(namen)
        if re.match(r"^(public|global)\b", s, re.I):
            oeffentlich.update(namen)
    if aktuell is not None:
        prozeduren.append(tuple(aktuell))
    return modul_namen, oeffentlich, prozeduren


def usage_check(files):
    all_problems = []
    per_file = {}
    modul_namen, oeffentlich = {}, set()

    for p in files:
        pr, _decl, lines = check(p)
        all_problems += [f"[{p}] {x}" for x in pr]
        per_file[p] = lines
        mn, oe, proz = zerlege(lines)
        modul_namen[p] = (mn, proz)
        oeffentlich |= oe

    # --- verbotene Aufrufe auf Auflistungen ------------------------
    VERBOTEN = [
        (r"\bSelectedSheets\s*\.\s*ExportAsFixedFormat",
         "SelectedSheets ist eine Sheets-Auflistung; ExportAsFixedFormat "
         "gibt es nur auf Workbook, Worksheet und Range"),
        (r"\bSelectedSheets\s*\.\s*PrintOut",
         "PrintOut auf SelectedSheets ist unzuverlaessig - ueber die Mappe gehen"),
        (r"\bSheets\s*\(\s*Array\(.*?\)\s*\)\s*\.\s*ExportAsFixedFormat",
         "ExportAsFixedFormat gibt es auf einer Sheets-Auflistung nicht"),
    ]
    for p, lines in per_file.items():
        for n, l in lines:
            for pat, hinweis in VERBOTEN:
                if re.search(pat, l, re.I):
                    all_problems.append(f"[{p}] Zeile {n}: {hinweis}")

    # --- Bezeichner JE PROZEDUR ------------------------------------
    # Frueher lief das ueber alle Dateien gemeinsam: eine Variable, die
    # irgendwo anders deklariert war, hat jede Verwendung durchgewinkt.
    # Genau so sind "Variable nicht definiert"-Fehler bis zum Nutzer
    # durchgerutscht.
    for p, (mn, proz) in modul_namen.items():
        erlaubt_modul = mn | oeffentlich
        for nm, startzeile, eigene, body in proz:
            erlaubt = erlaubt_modul | eigene
            for n, l in body:
                s = l.strip()
                if not s or s.lower().startswith("attribute"):
                    continue
                nostr = re.sub(r'"[^"]*"', ' ', s)
                nostr = re.sub(r"(?<![.\w])[A-Za-z_]\w*\s*:=", " ", nostr)
                for tok in re.findall(r"(?<![.\w])([A-Za-z_]\w*)", nostr):
                    lt = tok.lower()
                    if lt in VBA_KEYWORDS or lt in BUILTIN or lt in erlaubt:
                        continue
                    if lt.startswith("vb"):
                        continue
                    if lt.startswith(("xl", "mso")):
                        if lt not in OFFICE_CONSTANTS:
                            all_problems.append(
                                f"[{p}] Zeile {n}: unbekannte Office-Konstante {tok!r}")
                        continue
                    if lt in MODULES:
                        continue
                    all_problems.append(
                        f"[{p}] Zeile {n}: {tok!r} ist in {nm} nicht deklariert "
                        f"(Option Explicit) - {s[:70]}")

    # --- Modulqualifizierte Aufrufe: modXxx.Name --------------------
    for p, lines in per_file.items():
        for n, l in lines:
            s = l.strip()
            if not s:
                continue
            nostr = re.sub(r'"[^"]*"', ' ', s)
            for m in re.finditer(r"(?<![.\w])(mod[A-Za-z]\w*)\.([A-Za-z_]\w*)", nostr):
                modname, member = m.group(1), m.group(2)
                if modname.lower() not in MODULES:
                    continue
                if member.lower() not in oeffentlich:
                    all_problems.append(
                        f"[{p}] Zeile {n}: {modname}.{member} gibt es nicht "
                        f"oder ist dort Private")
    return all_problems


# ----------------------------------------------------------- Selbsttest
#
#  Bis 05.09.2026 wurde jede neue Regel dieser Datei von Hand gegen einen
#  kuenstlich eingebauten Fehler gehalten - also nur beim Bauen, und nur
#  fuer die gerade neue Regel. Die aelteren Regeln hat seit ihrer
#  Entstehung nie wieder jemand nachgewiesen. Genau das ist der Fehler,
#  den der Excel-Selbsttest vorgemacht hat: "keine leeren Zeilen" war
#  gruen, waehrend 35 dastanden.
#
#  Hier steht dieselbe Gegenprobe als Code: jede Regelgruppe bekommt
#  einen Fehler eingebaut und MUSS anschlagen. Faellt eine Regel beim
#  Umbauen still aus, faellt das ab jetzt sofort auf.

def _erste_prozedur(text):
    for i, l in enumerate(text.split("\n")):
        if re.match(r"^\s*(public\s+|private\s+|friend\s+)?(static\s+)?"
                    r"(sub|function|property)\b", l, re.I):
            return i
    raise ValueError("keine Prozedur gefunden")


def _vor_erster_prozedur(text, einschub):
    zeilen = text.split("\n")
    zeilen.insert(_erste_prozedur(text), einschub)
    return "\n".join(zeilen)


def _nach_erstem_end_sub(text, einschub):
    zeilen = text.split("\n")
    for i, l in enumerate(zeilen):
        if re.match(r"^\s*end\s+sub\b", l, re.I):
            zeilen.insert(i + 1, einschub)
            return "\n".join(zeilen)
    raise ValueError("kein End Sub gefunden")


def _anhaengen(text, rumpf):
    return text.rstrip("\n") + "\n\n" + rumpf.strip("\n") + "\n"


MUTATIONEN = [
    # (Name, erwarteter Textbaustein in der Meldung, Umbau)
    ("Anweisung im Deklarationsteil", "Anweisung im Deklarationsteil",
     lambda t: _vor_erster_prozedur(t, 'Debug.Print "Mutation"')),

    ("Deklaration nach der ersten Prozedur",
     "modulweite Deklaration nach der ersten Prozedur",
     lambda t: _nach_erstem_end_sub(t, "Private Const MUT_X As Long = 1")),

    ("Block nicht geschlossen", "ohne Anfang",
     lambda t: _anhaengen(t, """
Public Sub Mut_Block()
    If 1 = 1 Then
End Sub
""")),

    ("Zugriff auf FormatConditions", "FormatConditions",
     lambda t: _anhaengen(t, """
Public Sub Mut_CF()
    Dim ws As Worksheet
    ws.Cells.FormatConditions.Delete
End Sub
""")),

    ("Bereich schneidet verbundene Ferienzeilen",
     "schneidet die verbundenen Ferienzeilen",
     lambda t: _anhaengen(t, """
Public Sub Mut_Merge()
    Dim ws As Worksheet
    Dim letzte As Long
    ws.Range("C" & WP_FIRST_ROW & ":M" & letzte).Interior.Color = 1
End Sub
""")),

    ("undeklarierter Bezeichner", "nicht deklariert",
     lambda t: _anhaengen(t, """
Public Sub Mut_Explicit()
    xyzGibtEsNichtVariable = 1
End Sub
""")),

    ("unbekannte Office-Konstante", "unbekannte Office-Konstante",
     lambda t: _anhaengen(t, """
Public Sub Mut_Konstante()
    Dim x As Long
    x = xlNoRestriction
End Sub
""")),

    ("Aufruf eines nicht vorhandenen Modulmitglieds", "gibt es nicht",
     lambda t: _anhaengen(t, """
Public Sub Mut_Modulaufruf()
    modWochenplan.DiesenNamenGibtEsNicht
End Sub
""")),

    ("verbotener Aufruf auf einer Auflistung",
     "ExportAsFixedFormat gibt es nur auf",
     lambda t: _anhaengen(t, """
Public Sub Mut_Verboten()
    ActiveWindow.SelectedSheets.ExportAsFixedFormat Type:=xlTypePDF
End Sub
""")),
]


def selbsttest(dateien):
    """Baut jeden Fehler ein und verlangt, dass die zustaendige Regel anschlaegt."""
    import os, shutil, tempfile

    if usage_check(dateien):
        print("Ausgangslage ist nicht sauber - Selbsttest waere wertlos.")
        for x in usage_check(dateien)[:5]:
            print("   " + x)
        return 1

    ziel = None
    for p in dateien:
        if os.path.basename(p).lower() == "modwochenplan.bas":
            ziel = p
    if ziel is None:
        ziel = dateien[0]
    print("Ausgangsbasis: %s (sauber)\n" % os.path.basename(ziel))

    urtext = open(ziel, encoding="cp1252").read()
    tmp = tempfile.mkdtemp(prefix="vbacheck_")
    schlecht = 0
    print("Mutationen (jede MUSS anschlagen):")
    for name, baustein, umbau in MUTATIONEN:
        try:
            kaputt = umbau(urtext)
        except Exception as e:
            print("   FEHLT %-42s -> Mutation nicht baubar: %r" % (name, e))
            schlecht += 1
            continue
        pfad = os.path.join(tmp, os.path.basename(ziel))
        with open(pfad, "w", encoding="cp1252", newline="") as f:
            f.write(kaputt)
        satz = [pfad] + [d for d in dateien if d != ziel]
        meldungen = usage_check(satz)
        if any(baustein in m for m in meldungen):
            print("   ok    %-42s -> %s" % (name, baustein))
        else:
            print("   FEHLT %-42s -> nicht gemeldet (%d andere Meldungen)"
                  % (name, len(meldungen)))
            for m in meldungen[:3]:
                print("            %s" % m[:110])
            schlecht += 1
    shutil.rmtree(tmp, ignore_errors=True)
    print("\n%d von %d Mutationen erkannt."
          % (len(MUTATIONEN) - schlecht, len(MUTATIONEN)))
    return 1 if schlecht else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--selbsttest"]
    if "--selbsttest" in sys.argv[1:]:
        sys.exit(selbsttest(args))
    probs = usage_check(args)
    if not probs:
        print("OK - keine Beanstandungen")
        sys.exit(0)
    seen = set()
    for x in probs:
        if x not in seen:
            seen.add(x)
            print(x)
            #  Ausserhalb der CI tut das nichts. In der CI haengt es die
            #  Meldung an die Codezeile, auf die sie sich bezieht.
            ci_ausgabe.annotiere(x)
    print(f"\n{len(seen)} Meldung(en)")
    sys.exit(1)
