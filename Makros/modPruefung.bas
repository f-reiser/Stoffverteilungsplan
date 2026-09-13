Attribute VB_Name = "modPruefung"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Analyse der Mappe
'  ------------------------------------------------------------------
'  Der Blattschutz haelt den Anwender von den Formelspalten fern,
'  aber er laesst sich abschalten - und danach zerlegt ein Sortieren,
'  ein Ausschneiden-und-Einfuegen oder eine geloeschte Spalte die
'  Bezuege reihum. Sichtbar wird das oft erst Wochen spaeter.
'  Dieses Modul sieht nach und MELDET.
'
'  ES REPARIERT NICHTS. Das ist keine Bequemlichkeit, sondern die
'  Vorgabe aus Issue #31: Der Anwender soll erst seine Daten sichern
'  koennen. Ein missglueckter Reparaturlauf kann schlimmer sein als
'  der gemeldete Zustand. Repariert wird nur auf Knopfdruck, mit
'  "Einrichtung / Reparatur".
'
'  Die bedingte Formatierung wird GEZAEHLT, nicht gelesen: beim
'  Zugriff auf eine einzelne Regel ist Excel am 31.08.2026 hart
'  abgestuerzt, deshalb bleibt es bei FormatConditions.Count. Fuer die
'  Frage, ob der Farbbereich noch alle Planzeilen erreicht, genuegt
'  das. Ob die Regeln inhaltlich noch greifen, wird ueber ihre
'  Voraussetzungen geprueft: die Hilfsspalte V (ueber die zwei Regeln
'  laufen), die Gueltigkeitsliste der Spalte K und die Frage, ob die
'  Werte in K ueberhaupt in der Stundenliste stehen - steht dort etwas
'  anderes, findet keine Farbregel etwas.
'=====================================================================

'  Die Kategorien, unter denen ein Befund gemeldet wird. Der
'  Selbsttest fragt ueber sie ab, der Anwender liest sie in der
'  Warnzeile - deshalb sind es Klartexte und keine Nummern.
Public Const PRF_FEHLERWERT    As String = "Fehlerwert"
Public Const PRF_FORMEL_FEHLT  As String = "Formel fehlt"
Public Const PRF_FORMEL_ANDERS As String = "Formel weicht ab"
Public Const PRF_QUERBEZUG     As String = "Querbezug"
Public Const PRF_GUELTIGKEIT   As String = "Gueltigkeitsliste"
Public Const PRF_KATEGORIE     As String = "Stunden-Kategorie"
Public Const PRF_FARBREGEL     As String = "Farbregeln"

'  Name der Zelle im Blatt "Steuerung", in der die Warnung steht.
'  Ueber einen Namen und nicht ueber eine Zeilennummer - im ganzen
'  Projekt ist keine Zeile fest verdrahtet.
Public Const PRF_ZELLE_NAME As String = "wpPruefHinweis"

Public Const PRF_SAUBER As String = _
    "Prüfung: keine Auffälligkeiten. Formeln, Bezüge, Farbregeln und die " & _
    "Auswahlliste der Spalte K sind in Ordnung."

'  Wie viele Beispiele je Befund genannt werden. Eine vollstaendige
'  Liste liest in einer Warnzeile ohnehin niemand, und bei einem
'  zerlegten Blatt waere sie hunderte Eintraege lang.
Private Const PRF_BEISPIELE As Long = 6

'  Spalten, in denen in JEDER Inhaltszeile dieselbe Formel stehen muss.
'  Bewusst ohne B (nach dem Fixieren feste Werte), ohne S (verweist auf
'  die vorherige Inhaltszeile, die unterschiedlich weit weg liegt) und
'  ohne U (eigene Pruefung, siehe Querbezug).
Private Const PRF_WP_SPALTEN As String = "A,C,D,L,P,Q,R,V"
Private Const PRF_LB_SPALTEN As String = "F,G,H,I"


'=====================================================================
'  Alle Befunde, je Zeile "<Kategorie>: <Text>". Leer = nichts gefunden.
'=====================================================================
Public Function MappePruefen() As String
    Dim s As String
    s = s & Fehlerwerte()
    s = s & Planformeln()
    s = s & Lernbereichsformeln()
    s = s & Querbezug()
    s = s & Gueltigkeit()
    s = s & Farbregeln()
    MappePruefen = s
End Function


'=====================================================================
'  Nur die Befunde einer Kategorie, aus einem fertigen MappePruefen.
'  Leer = diese Pruefung ist gruen.
'
'  Der Aufrufer reicht das Ergebnis herein, statt dass hier neu geprueft
'  wird: der Selbsttest fragt sieben Kategorien nacheinander ab, und
'  eine Analyse liest ueber alle Planzeilen hinweg mehrere hundert
'  Zellen einzeln. Siebenmal von vorn waere im Mutationstest, der den
'  Abschnitt gut dreissigmal laufen laesst, deutlich spuerbar.
'=====================================================================
Public Function BefundAus(ByVal alleBefunde As String, ByVal kat As String) As String
    Dim alle As Variant, i As Long, s As String, z As String

    alle = Split(alleBefunde, vbCrLf)
    For i = LBound(alle) To UBound(alle)
        z = CStr(alle(i))
        If Left$(z, Len(kat) + 1) = kat & ":" Then
            If Len(s) > 0 Then s = s & vbCrLf
            s = s & z
        End If
    Next i
    BefundAus = s
End Function


'=====================================================================
'  Die Warnung anzeigen: Zeile im Blatt "Steuerung" und Statusleiste.
'  Aufgerufen von Setup_Stoffverteilungsplan und von Workbook_Open.
'=====================================================================
Public Sub WarnungAnzeigen()
    Dim c As Range, ws As Worksheet, txt As String, neu As String, n As Long

    On Error GoTo Fail
    txt = MappePruefen()
    n = Befundzahl(txt)

    If n = 0 Then
        neu = PRF_SAUBER
    Else
        neu = ChrW$(9888) & " " & n & " Auffälligkeit(en) gefunden: " & Einzeiler(txt) & _
              " - Bitte erst eine Sicherungskopie anlegen, dann """ & CAP_SETUP & """."
    End If

    Set c = HinweisZelle()
    If Not c Is Nothing Then
        '  Nur schreiben, wenn sich wirklich etwas geaendert hat: jeder
        '  Schreibzugriff macht die Mappe "geaendert", und dann fragt
        '  Excel beim Schliessen nach dem Speichern, obwohl der Anwender
        '  die Datei nur aufgemacht hat.
        If CStr(c.Value) <> neu Then
            Set ws = c.Worksheet
            modSchutz.Schutz_Aus ws
            c.Value = neu
            c.Font.Bold = (n > 0)
            If n > 0 Then
                c.Font.Color = RGB(186, 74, 74)
            Else
                c.Font.Color = modWochenplan.FARBE_LEISE
            End If
            modSchutz.Schutz_An
        End If
    End If

    StatusleisteSetzen n
    Exit Sub
Fail:
    modSchutz.Schutz_An
    StatusleisteSetzen 0
End Sub


'  Die Statusleiste ist ausdruecklich nur die Zugabe (Issue #31:
'  "falls moeglich"). False gibt sie an Excel zurueck - ein
'  stehengebliebener eigener Text wuerde sonst jede spaetere
'  Excel-Meldung ueberdecken.
Private Sub StatusleisteSetzen(ByVal n As Long)
    On Error Resume Next
    If n = 0 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = ChrW$(9888) & " Stoffverteilungsplan: " & n & _
            " Auffälligkeit(en) - siehe Blatt """ & CTRL_SHEET & """."
    End If
    On Error GoTo 0
End Sub


Private Function HinweisZelle() As Range
    On Error Resume Next
    Set HinweisZelle = ThisWorkbook.Names(PRF_ZELLE_NAME).RefersToRange
    On Error GoTo 0
End Function


Private Function Befundzahl(ByVal txt As String) As Long
    Dim teile As Variant, i As Long, n As Long
    teile = Split(txt, vbCrLf)
    For i = LBound(teile) To UBound(teile)
        If Len(Trim$(CStr(teile(i)))) > 0 Then n = n + 1
    Next i
    Befundzahl = n
End Function


Private Function Einzeiler(ByVal txt As String) As String
    Dim teile As Variant, i As Long, s As String
    teile = Split(txt, vbCrLf)
    For i = LBound(teile) To UBound(teile)
        If Len(Trim$(CStr(teile(i)))) > 0 Then
            If Len(s) > 0 Then s = s & " | "
            s = s & Trim$(CStr(teile(i)))
        End If
    Next i
    Einzeiler = Kurz(s, 400)
End Function


'=====================================================================
'  1. Fehlerwerte in den Zellen
'  ------------------------------------------------------------------
'  #BEZUG! ist der Fall aus Issue #31: ein Bezug, der ins Leere zeigt.
'  Gelesen wird der ganze UsedRange in EINEM Zugriff - zellweise waere
'  das ueber drei Blaetter die haeufigste Bremse.
'=====================================================================
Private Function Fehlerwerte() As String
    Dim namen As Variant, i As Long, s As String
    namen = Array(WP_SHEET, LB_SHEET, SET_SHEET)
    For i = LBound(namen) To UBound(namen)
        s = s & FehlerwerteImBlatt(CStr(namen(i)))
    Next i
    Fehlerwerte = s
End Function


Private Function FehlerwerteImBlatt(ByVal nm As String) As String
    Dim ws As Worksheet, rng As Range, v As Variant
    Dim r As Long, c As Long, r0 As Long, c0 As Long
    Dim n As Long, liste As String, zelle As Range, frm As String

    Set ws = BlattOderNichts(nm)
    If ws Is Nothing Then Exit Function

    Set rng = ws.UsedRange
    r0 = rng.Row
    c0 = rng.Column
    '  Ein einzelliger Bereich liefert keinen Array, sondern den Wert
    '  selbst - dann gaebe es unten zwei Wege durch dieselbe Schleife.
    If rng.Rows.Count = 1 And rng.Columns.Count = 1 Then Set rng = rng.Resize(2, 2)
    v = rng.Value

    For r = LBound(v, 1) To UBound(v, 1)
        For c = LBound(v, 2) To UBound(v, 2)
            If IsError(v(r, c)) Then
                n = n + 1
                If n <= PRF_BEISPIELE Then
                    Set zelle = ws.Cells(r0 + r - 1, c0 + c - 1)
                    frm = ""
                    On Error Resume Next
                    frm = zelle.Formula
                    On Error GoTo 0
                    If Len(liste) > 0 Then liste = liste & ", "
                    liste = liste & zelle.Address(False, False)
                    If InStr(1, frm, "#REF!", vbTextCompare) > 0 Then
                        liste = liste & " (Bezug zeigt ins Leere)"
                    ElseIf Len(frm) > 0 Then
                        liste = liste & " (" & Kurz(frm, 40) & ")"
                    End If
                End If
            End If
        Next c
    Next r

    If n > 0 Then
        FehlerwerteImBlatt = Meldung(PRF_FEHLERWERT, _
            "Blatt """ & nm & """: " & n & " Zelle(n) mit Fehlerwert - " & liste)
    End If
End Function


'=====================================================================
'  2. Formeln des Wochenplans
'  ------------------------------------------------------------------
'  Verglichen wird ueber FormulaR1C1, nicht ueber Formula. In der
'  R1C1-Schreibweise steht ein zeilenrelativer Bezug in JEDER Zeile
'  zeichengleich da ("RC5" statt "$E4", "$E5", ...) - zwei Zeilen sind
'  damit ohne jede Normierung vergleichbar.
'
'  Die Frage ist bewusst nicht "stimmt die Formel mit dem ueberein, was
'  modWochenplan schreiben wuerde": Test und Code kaemen dann aus
'  derselben Vorstellung und haetten denselben blinden Fleck. Gefragt
'  wird stattdessen "sehen alle Planzeilen gleich aus" - eine Zeile,
'  die aus der Reihe faellt, ist immer ein Fehler, ganz gleich welche
'  Formel die richtige waere.
'=====================================================================
Private Function Planformeln() As String
    Dim ws As Worksheet, letzte As Long, spalten As Variant, i As Long, s As String

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Function
    letzte = modWochenplan.PlanLastRow(ws)
    spalten = Split(PRF_WP_SPALTEN, ",")
    For i = LBound(spalten) To UBound(spalten)
        s = s & SpaltePruefen(ws, WP_SHEET, CStr(spalten(i)), _
                              WP_FIRST_ROW, letzte, True)
    Next i
    Planformeln = s
End Function


Private Function Lernbereichsformeln() As String
    Dim zl As Worksheet, z1 As Long, zSum As Long, zLetzte As Long
    Dim spalten As Variant, i As Long, s As String

    Set zl = modWochenplan.LbSheet()
    If zl Is Nothing Then Exit Function
    z1 = modWochenplan.LB_FIRST_ROW()
    zSum = modWochenplan.LbSummeRow(zl)
    zLetzte = modWochenplan.LbLastDataRow(zl, zSum)
    If zLetzte < z1 Then Exit Function

    spalten = Split(PRF_LB_SPALTEN, ",")
    For i = LBound(spalten) To UBound(spalten)
        s = s & SpaltePruefen(zl, LB_SHEET, CStr(spalten(i)), z1, zLetzte, False)
    Next i
    Lernbereichsformeln = s
End Function


Private Function SpaltePruefen(ByVal ws As Worksheet, ByVal blattName As String, _
                               ByVal col As String, ByVal r1 As Long, _
                               ByVal r2 As Long, ByVal ohneFerien As Boolean) As String
    Dim feld As Variant, r As Long, nimm As Boolean, z As String
    Dim muster As String, musterZeile As Long
    Dim fehlt As String, anders As String, nFehlt As Long, s As String

    If r2 < r1 Then Exit Function
    feld = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).FormulaR1C1

    For r = r1 To r2
        nimm = True
        If ohneFerien Then
            If modWochenplan.IsFerienRow(ws, r) Then nimm = False
        End If
        If nimm Then
            z = FormelAus(feld, r - r1 + 1)
            If Left$(z, 1) <> "=" Then
                nFehlt = nFehlt + 1
                If nFehlt <= PRF_BEISPIELE Then fehlt = fehlt & " " & col & r
            ElseIf Len(muster) = 0 Then
                muster = z
                musterZeile = r
            ElseIf z <> muster And Len(anders) = 0 Then
                anders = col & r & " (Vergleichszeile " & musterZeile & ")"
            End If
        End If
    Next r

    If Len(fehlt) > 0 Then
        s = s & Meldung(PRF_FORMEL_FEHLT, "Blatt """ & blattName & """: " & nFehlt & _
                        " Zelle(n) ohne Formel -" & fehlt)
    End If
    If Len(anders) > 0 Then
        s = s & Meldung(PRF_FORMEL_ANDERS, "Blatt """ & blattName & """: " & anders)
    End If
    SpaltePruefen = s
End Function


'  Eine Spalte mit nur einer Zeile liefert keinen Array.
Private Function FormelAus(ByRef feld As Variant, ByVal i As Long) As String
    If IsArray(feld) Then
        FormelAus = CStr(feld(i, 1))
    ElseIf i = 1 Then
        FormelAus = CStr(feld)
    End If
End Function


'=====================================================================
'  3. Querbezug der Spalte U auf die Lernbereiche
'  ------------------------------------------------------------------
'  Zeigt der FILTER noch auf den Stand vor einer Verschiebung, bleibt
'  die Spalte still LEER - ohne Fehlermeldung, ohne #BEZUG!. Genau
'  dieser Fehler steckte einmal drin ($A$2:$A$50 statt $A$4:$A$52),
'  und er ist an einer Formel mit sich selbst nicht zu erkennen:
'  alle Zeilen sind gleich falsch.
'=====================================================================
Private Function Querbezug() As String
    Dim ws As Worksheet, r As Long, letzte As Long, soll As String, frm As String

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Function
    If modWochenplan.LbSheet() Is Nothing Then Exit Function

    soll = LB_SHEET & "!$A$" & modWochenplan.LB_FIRST_ROW()
    letzte = modWochenplan.PlanLastRow(ws)

    For r = WP_FIRST_ROW To letzte
        If Not modWochenplan.IsFerienRow(ws, r) Then
            frm = ""
            On Error Resume Next
            frm = ws.Cells(r, "U").Formula
            On Error GoTo 0
            If Left$(frm, 1) <> "=" Then
                Querbezug = Meldung(PRF_QUERBEZUG, _
                    "Spalte U hat in Zeile " & r & " keine Formel - der Abgleich " & _
                    "mit den Lernbereichen findet nicht statt.")
                Exit Function
            ElseIf InStr(1, frm, soll, vbTextCompare) = 0 Then
                Querbezug = Meldung(PRF_QUERBEZUG, _
                    "Spalte U sucht die Lernbereiche nicht bei " & soll & _
                    " (Zeile " & r & "): " & Kurz(frm, 70))
                Exit Function
            End If
        End If
    Next r
End Function


'=====================================================================
'  4. Auswahlliste der Spalte K
'  ------------------------------------------------------------------
'  An der Stundenliste haengen zwei Dinge: die Gueltigkeitsliste der
'  Spalte K und fuenf Regeln der bedingten Formatierung, die den Wert
'  in K in der Liste SUCHEN. Steht in K etwas, das dort nicht
'  vorkommt, findet keine Regel etwas - die Zeile bleibt farblos, und
'  eine Fehlermeldung gibt es nicht.
'=====================================================================
Private Function Gueltigkeit() As String
    Dim ws As Worksheet, st As Worksheet, r As Long, letzte As Long
    Dim adr As String, soll As String, ist As String, p As Long
    Dim liste As Variant, wert As String
    Dim ohne As String, nOhne As Long, falsch As String
    Dim fremd As String, nFremd As Long, s As String

    Set ws = modWochenplan.WpSheet()
    Set st = modWochenplan.SetSheet()
    If ws Is Nothing Or st Is Nothing Then Exit Function

    adr = modWochenplan.StdListAddress()
    p = InStr(adr, "!")
    If p = 0 Then Exit Function
    liste = st.Range(Mid$(adr, p + 1)).Value
    soll = NormAdr(adr)

    letzte = modWochenplan.PlanLastRow(ws)
    For r = WP_FIRST_ROW To letzte
        If Not modWochenplan.IsFerienRow(ws, r) Then
            ist = ListenAdresse(ws.Cells(r, "K"))
            If Len(ist) = 0 Then
                nOhne = nOhne + 1
                If nOhne <= PRF_BEISPIELE Then ohne = ohne & " K" & r
            ElseIf NormAdr(ist) <> soll And Len(falsch) = 0 Then
                falsch = "K" & r & " zeigt auf " & ist
            End If

            wert = ZellText(ws.Cells(r, "K"))
            If Len(wert) > 0 Then
                If Not InListe(liste, wert) Then
                    nFremd = nFremd + 1
                    If nFremd <= PRF_BEISPIELE Then fremd = fremd & " K" & r & "=""" & wert & """"
                End If
            End If
        End If
    Next r

    If nOhne > 0 Then
        s = s & Meldung(PRF_GUELTIGKEIT, nOhne & " Zelle(n) der Spalte K ohne " & _
                        "Auswahlliste -" & ohne)
    End If
    If Len(falsch) > 0 Then
        s = s & Meldung(PRF_GUELTIGKEIT, "Die Auswahlliste der Spalte K zeigt nicht " & _
                        "auf die Stundenliste " & adr & ": " & falsch)
    End If
    If nFremd > 0 Then
        s = s & Meldung(PRF_KATEGORIE, nFremd & " Wert(e) in Spalte K stehen nicht in " & _
                        "der Stundenliste " & adr & " - dort greift keine Farbregel:" & fremd)
    End If
    Gueltigkeit = s
End Function


'  Die Liste, auf die die Gueltigkeit einer Zelle zeigt. Leer heisst
'  "keine Gueltigkeit oder keine Liste" - Validation.Type wirft einen
'  Fehler, wenn gar keine da ist, und das ist hier ein Ergebnis und
'  keine Stoerung.
Private Function ListenAdresse(ByVal c As Range) As String
    Dim art As Long, f As String

    art = 0
    On Error Resume Next
    art = c.Validation.Type
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    f = c.Validation.Formula1
    On Error GoTo 0
    If art <> xlValidateList Then Exit Function
    ListenAdresse = f
End Function


'  Adressen vergleichbar machen: Excel liefert die Formel einer
'  Gueltigkeit mit fuehrendem "=", Blattnamen je nach Laune in
'  Hochkommata.
Private Function NormAdr(ByVal s As String) As String
    NormAdr = UCase$(Replace(Replace(Replace(s, "=", ""), "'", ""), " ", ""))
End Function


Private Function InListe(ByRef liste As Variant, ByVal wert As String) As Boolean
    Dim i As Long, j As Long

    If IsArray(liste) Then
        For i = LBound(liste, 1) To UBound(liste, 1)
            For j = LBound(liste, 2) To UBound(liste, 2)
                If Not IsError(liste(i, j)) Then
                    If StrComp(Trim$(CStr(liste(i, j) & "")), wert, vbTextCompare) = 0 Then
                        InListe = True
                        Exit Function
                    End If
                End If
            Next j
        Next i
    ElseIf Not IsError(liste) Then
        InListe = (StrComp(Trim$(CStr(liste & "")), wert, vbTextCompare) = 0)
    End If
End Function


'=====================================================================
'  5. Reichweite der Farbregeln
'  ------------------------------------------------------------------
'  Acht Regeln faerben den Wochenplan. Erwartet wird hier trotzdem
'  KEINE feste Regelzahl: sieben der acht sind x14-Erweiterungsregeln,
'  und ob VBA die mitzaehlt, ist nirgends verbrieft. Eine fest
'  eingetragene Acht koennte deshalb auf einer voellig gesunden Mappe
'  bei jedem Oeffnen Alarm schlagen.
'
'  Verglichen werden stattdessen die Planzeilen untereinander: eine
'  Zeile, die weniger Regeln traegt als die erste, ist aus dem
'  Farbbereich herausgewachsen. Das ist unabhaengig davon, welche Zahl
'  Excel meldet - und es ist genau der Fall, der beim Anwender
'  auftritt, wenn er unten Zeilen anhaengt oder oben welche loescht.
'=====================================================================
Private Function Farbregeln() As String
    Dim ws As Worksheet, r As Long, letzte As Long
    Dim soll As Long, ist As Long, n As Long, liste As String
    Dim erg As String, flaechen As Long

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Function
    letzte = modWochenplan.PlanLastRow(ws)

    '  Gibt Excel die Zahl gar nicht her, wird geschwiegen. Eine 0 als
    '  Ersatzwert waere hier die schlimmere Antwort: sie ergaebe bei
    '  jedem Oeffnen eine Warnung ueber eine voellig heile Mappe.
    soll = RegelZahl(ws, WP_FIRST_ROW)
    If soll < 0 Then Exit Function
    If soll = 0 Then
        Farbregeln = Meldung(PRF_FARBREGEL, "Der Wochenplan hat ab Zeile " & _
            WP_FIRST_ROW & " keine bedingte Formatierung mehr - Kategorie, " & _
            "halbe Klasse und halbe Woche bleiben farblos.")
        Exit Function
    End If

    For r = WP_FIRST_ROW + 1 To letzte
        ist = RegelZahl(ws, r)
        If ist >= 0 And ist <> soll Then
            n = n + 1
            If n <= PRF_BEISPIELE Then
                If Len(liste) > 0 Then liste = liste & ", "
                liste = liste & "Zeile " & r & " (" & ist & " statt " & soll & ")"
            End If
        End If
    Next r

    If n > 0 Then
        erg = erg & Meldung(PRF_FARBREGEL, n & " Planzeile(n) tragen andere " & _
            "Farbregeln als Zeile " & WP_FIRST_ROW & ": " & liste)
    End If

    '  Die Regelzahl bleibt gleich, wenn eine Regel nicht verschwindet,
    '  sondern nur zerfaellt: wiederholtes Einfuegen/Loeschen zerlegt
    '  ihre Flaeche ("Wird angewendet auf") in mehrere Teilbereiche, die
    '  zusammen weiterhin jede Zeile abdecken (siehe PR #59). Das sieht
    '  keine Zaehlung - deshalb der zusaetzliche Zugriff unten.
    flaechen = FlaechenZerfallen(ws)
    If flaechen > 1 Then
        erg = erg & Meldung(PRF_FARBREGEL, "Die Farbregeln sind in " & flaechen & _
            " Teilbereiche zerfallen, obwohl die Regelzahl je Zeile noch stimmt - " & _
            "vermutlich durch wiederholtes Einfuegen/Loeschen von Zeilen. Die " & _
            "Formatierung ist dann unzuverlaessig.")
    End If

    Farbregeln = erg
End Function


'  Wieviele Teilbereiche die erste Regel abdeckt (siehe Farbregeln).
'  >1 heisst: die Flaeche ist zerfallen. -1 heisst wie bei RegelZahl
'  "Excel hat nichts hergegeben".
'
'  Das ist der EINE zusaetzliche Zugriff ueber .Count hinaus, den
'  Regel 2 aus CLAUDE.md seit dem 13.09.2026 erlaubt (PR #59, "Weg 2"):
'  ein fester, literaler Index (1), keine Schleife ueber die Regeln,
'  nur AppliesTo.Areas.Count - kein Formula1, kein Schreiben.
'  vbacheck.py haelt genau dieses eine Muster nach.
'
'  Faellt es Excel wieder hart um: nur diese Funktion und ihr Aufruf
'  in Farbregeln muessen zurueckgebaut werden, dazu die Lockerung in
'  vbacheck.py (Abschnitt "3. FormatConditions").
Private Function FlaechenZerfallen(ByVal ws As Worksheet) As Long
    FlaechenZerfallen = -1
    On Error Resume Next
    FlaechenZerfallen = ws.Cells(WP_FIRST_ROW, "B").FormatConditions(1).AppliesTo.Areas.Count
    On Error GoTo 0
End Function


'  Wie viele Regeln der bedingten Formatierung auf diese Zeile wirken.
'  Gefragt wird in Spalte B: dort ist auch eine ueber B:M verbundene
'  Ferienzeile verankert, und die Farbregeln beginnen ohnehin dort.
'
'  Nur .Count - kein Zugriff auf eine einzelne Regel. vbacheck.py
'  haelt das nach, siehe den Kopf dieses Moduls.
'
'  -1 heisst "Excel hat die Zahl nicht hergegeben" und ist etwas
'  anderes als 0 ("keine Regel da") - siehe Farbregeln.
Private Function RegelZahl(ByVal ws As Worksheet, ByVal r As Long) As Long
    RegelZahl = -1
    On Error Resume Next
    RegelZahl = ws.Cells(r, "B").FormatConditions.Count
    On Error GoTo 0
End Function


'=====================================================================
'  Kleinkram
'=====================================================================
Private Function Meldung(ByVal kat As String, ByVal txt As String) As String
    Meldung = kat & ": " & txt & vbCrLf
End Function


Private Function Kurz(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then
        Kurz = s
    Else
        Kurz = Left$(s, n - 3) & "..."
    End If
End Function


'  Der Wert einer Zelle als Text. Ein Fehlerwert laesst sich nicht in
'  einen String wandeln - dann gilt die Zelle hier als leer, gemeldet
'  wird sie ohnehin schon von Fehlerwerte().
Private Function ZellText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsError(v) Then Exit Function
    ZellText = Trim$(CStr(v & ""))
End Function


Private Function BlattOderNichts(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set BlattOderNichts = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function
