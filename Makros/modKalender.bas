Attribute VB_Name = "modKalender"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Kalender: Unterrichtswochen und Ferienzeilen
'  ------------------------------------------------------------------
'  Oeffentliche Makros:
'    UW_Und_Ferien_Generieren   Spalte E (UW) im Wochenplan als Werte
'                               neu schreiben und die Ferienzeilen neu
'                               aufbauen
'    Ferien_Aus_Kalender_Vorschlagen
'                               Ferientabelle in "Einstellungen" aus
'                               den Luecken der Schulwochen-Tabelle
'                               befuellen
'    Einstellungen_Erweitern    Fixierungs-Schalter und Ferientabelle
'                               im Blatt "Einstellungen" anlegen bzw.
'                               an die richtige Stelle verschieben
'
'  Ablauf von UW_Und_Ferien_Generieren:
'    1. alle vorhandenen Ferienzeilen entfernen
'    2. den verbleibenden Planzeilen der Reihe nach die verfuegbaren
'       Schulwochen (Einstellungen: "Woche verfuegbar" = WAHR) als
'       feste Werte in Spalte E zuweisen
'    3. fuer jeden Eintrag der Ferientabelle an der passenden Stelle
'       eine Ferienzeile einfuegen (B:M verbunden, eigene Formatierung,
'       Kennzeichen "F" in Spalte N)
'    4. Formeln und Gueltigkeitsliste neu aufbauen
'
'  Die Ferientabelle wird rein ueber die Datumsangaben ausgewertet -
'  es koennen dort also auch andere unterrichtsfreie Bloecke stehen
'  (z. B. "Betriebspraktikum").
'
'  Alle Zeilennummern im Blatt "Einstellungen" kommen aus
'  modWochenplan.RefreshLayout - hier ist nichts fest verdrahtet.
'=====================================================================

Private Const FER_FILL   As Long = 15132390     ' RGB(230, 230, 230)
Private Const FER_FONT   As Long = 6579300      ' RGB(100, 100, 100)
Private Const FER_HEIGHT As Double = 19.5

' Standardspalte fuer den Block "Angaben fuer den Seitenkopf" (O)
Private Const KOPF_DEF_COL As Long = 15


'=====================================================================
'  Hauptmakro
'=====================================================================
Public Sub UW_Und_Ferien_Generieren()
    Dim ws As Worksheet, st As Worksheet
    Dim wkNr() As Long, wkVon() As Date, wkBis() As Date
    Dim nWeek As Long
    Dim ferVon() As Date, ferBis() As Date, ferNam() As String, nFer As Long
    Dim lastRow As Long, nPlan As Long, i As Long, r As Long
    Dim startYear As Long
    Dim nSet As Long, nMiss As Long, nIns As Long, nSkip As Long, nLeer As Long
    Dim msg As String

    Set ws = modWochenplan.WpSheet()
    Set st = modWochenplan.SetSheet()
    If ws Is Nothing Or st Is Nothing Then
        modWochenplan.Info "Die Blätter '" & WP_SHEET & "' und '" & SET_SHEET & "' werden benötigt.", _
               vbExclamation, "Stoffverteilungsplan"
        Exit Sub
    End If
    If Not modWochenplan.LayoutReady() Then Exit Sub

    startYear = SchoolStartYear(st)
    nWeek = ReadWeeks(st, startYear, wkNr, wkVon, wkBis)
    If nWeek = 0 Then
        modWochenplan.Info "In der Tabelle 'Einstellungen' ist keine verfügbare Schulwoche eingetragen.", _
               vbExclamation, "Stoffverteilungsplan"
        Exit Sub
    End If
    nFer = ReadFerien(st, startYear, ferVon, ferBis, ferNam)

    modWochenplan.FastOn ws
    On Error GoTo Fail

    ' --- 1. vorhandene Ferienzeilen entfernen -------------------------
    modWochenplan.SetStep "Alte Ferienzeilen entfernen"
    lastRow = modWochenplan.PlanLastRow(ws)
    For r = lastRow To WP_FIRST_ROW Step -1
        If modWochenplan.IsFerienRow(ws, r) Then
            ws.Range(ws.Cells(r, "B"), ws.Cells(r, "M")).UnMerge
            ws.Rows(r).Delete Shift:=xlUp
        End If
    Next r

    ' --- 2. Unterrichtswochen als Werte schreiben ---------------------
    modWochenplan.SetStep "Unterrichtswochen schreiben"
    lastRow = modWochenplan.PlanLastRow(ws)
    nPlan = lastRow - WP_FIRST_ROW + 1
    For i = 1 To nPlan
        r = WP_FIRST_ROW + i - 1
        If i <= nWeek Then
            ws.Cells(r, "E").Value = wkNr(i)
            nSet = nSet + 1
        Else
            ws.Cells(r, "E").ClearContents
            nMiss = nMiss + 1
        End If
    Next i

    ' --- 2b. leere Zeilen am Ende entfernen ---------------------------
    ' Erst jetzt, nachdem die Unterrichtswochen verteilt sind: Zeilen
    ' ohne jeden Inhalt am Tabellenende sollen gar nicht erst
    ' stehenbleiben.
    modWochenplan.SetStep "Leerzeilen am Ende entfernen"
    nLeer = modWochenplan.LeereEndzeilenEntfernen(ws)

    ' --- 3. Ferienzeilen einfuegen (von hinten nach vorne) ------------
    modWochenplan.SetStep "Ferienzeilen einfuegen"
    For i = nFer To 1 Step -1
        If InsertFerienRow(ws, st, startYear, ferVon(i), ferBis(i), ferNam(i)) Then
            nIns = nIns + 1
        Else
            nSkip = nSkip + 1
        End If
    Next i

    ' --- 4. Formeln neu aufbauen --------------------------------------
    modWochenplan.SetStep "Formeln aufbauen"
    modWochenplan.RebuildAll ws, modWochenplan.PlanLastRow(ws)

    modWochenplan.FastOff

    msg = nSet & " Planzeilen haben eine Unterrichtswoche bekommen." & vbCrLf
    If nWeek > nSet Then
        msg = msg & "Achtung: " & (nWeek - nSet) & " verfügbare Schulwoche(n) sind " & _
              "nicht verplant (zu wenige Planzeilen)." & vbCrLf
    End If
    If nMiss > 0 Then
        msg = msg & "Achtung: " & nMiss & " Planzeile(n) blieben ohne Unterrichtswoche " & _
              "(zu wenige verfügbare Wochen)." & vbCrLf
    End If
    msg = msg & nIns & " Ferienzeile(n) eingefügt."
    If nLeer > 0 Then
        msg = msg & vbCrLf & nLeer & " leere Zeile(n) am Ende entfernt."
    End If
    If nSkip > 0 Then
        msg = msg & vbCrLf & nSkip & " Ferieneintrag/-einträge liegen außerhalb des " & _
              "verplanten Zeitraums und wurden übersprungen."
    End If

    ' Bewusst NICHT zurueck zur Steuerung: nach dem Erzeugen will man
    ' das Ergebnis sehen. Der Selbsttest prueft genau diesen Sprung.
    modWochenplan.GotoSheet WP_SHEET, "F" & WP_FIRST_ROW
    modWochenplan.Info msg
    Exit Sub
Fail:
    modWochenplan.ReportError "Erzeugen der Unterrichtswochen"
End Sub


'=====================================================================
'  Eine Ferienzeile an der passenden Stelle einfuegen
'  Rueckgabe: True, wenn eine Zeile entstanden ist
'=====================================================================
Private Function InsertFerienRow(ByVal ws As Worksheet, ByVal st As Worksheet, _
                                 ByVal startYear As Long, ByVal dVon As Date, _
                                 ByVal dBis As Date, ByVal nam As String) As Boolean
    Dim lastRow As Long, r As Long, t As Long
    Dim uw As Long, wVon As Date
    Dim txt As String

    lastRow = modWochenplan.PlanLastRow(ws)

    ' erste Planzeile suchen, deren Schulwoche NACH dem Ferienende beginnt
    t = 0
    For r = WP_FIRST_ROW To lastRow
        If Not modWochenplan.IsFerienRow(ws, r) Then
            If Len(Trim$(CStr(ws.Cells(r, "E").Value))) > 0 Then
                If IsNumeric(ws.Cells(r, "E").Value) Then
                    uw = CLng(ws.Cells(r, "E").Value)
                    If WeekStart(st, startYear, uw, wVon) Then
                        If wVon > dBis Then
                            t = r
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If t = 0 Then Exit Function

    ' Zeile t kopieren und als "kopierte Zellen" oberhalb einfuegen -
    ' dabei dehnt Excel die bedingte Formatierung selbst mit aus.
    Application.CutCopyMode = False
    ws.Rows(t).Copy
    ws.Rows(t).Insert Shift:=xlDown
    Application.CutCopyMode = False

    ws.Range(ws.Cells(t, "A"), ws.Cells(t, "V")).ClearContents
    modWochenplan.MarkierungSetzen ws, t

    txt = nam
    If Len(Trim$(txt)) = 0 Then txt = "unterrichtsfrei"
    txt = txt & "   " & Format$(dVon, "dd.mm.yyyy") & " - " & Format$(dBis, "dd.mm.yyyy")

    FormatFerienRow ws, t
    ws.Cells(t, "B").Value = txt

    InsertFerienRow = True
End Function


Private Sub FormatFerienRow(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    ' Format der Spalte G auf J uebertragen: damit verschwindet das
    ' Kontrollkaestchen "Erledigt" aus der Ferienzeile, bevor B:M
    ' verbunden wird.
    Application.CutCopyMode = False
    ws.Cells(r, "G").Copy
    ws.Cells(r, "J").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    With ws.Range(ws.Cells(r, "B"), ws.Cells(r, "M"))
        .UnMerge
        .ClearContents
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Bold = True
        .Font.Italic = True
        .Font.Color = FER_FONT
        .Interior.Pattern = xlSolid
        .Interior.Color = FER_FILL
        ' Ferienzeilen sind reine Anzeige - sie bleiben auch bei
        ' aktivem Blattschutz gesperrt.
        .Locked = True
    End With
    ws.Rows(r).RowHeight = FER_HEIGHT

    ' Zum Schluss, damit die Schriftfarbe zum endgueltigen Hintergrund
    ' der Zelle passt.
    modWochenplan.MarkierungSetzen ws, r
    On Error GoTo 0
End Sub


'=====================================================================
'  Einstellungen lesen
'=====================================================================
Public Function SchoolStartYear(ByVal st As Worksheet) As Long
    Dim v As Variant
    On Error Resume Next
    v = st.Cells(modWochenplan.YearRow(), 2).Value
    On Error GoTo 0
    If IsNumeric(v) Then
        SchoolStartYear = CLng(v)
    Else
        SchoolStartYear = Year(Date)
    End If
End Function


'  Datum auf das richtige Schuljahr normieren: Monate ab September
'  gehoeren zum Startjahr, alle uebrigen zum Folgejahr.
Private Function NormDate(ByVal d As Date, ByVal startYear As Long) As Date
    Dim y As Long
    If Month(d) >= 9 Then y = startYear Else y = startYear + 1
    NormDate = DateSerial(y, Month(d), Day(d))
End Function


'  Liefert alle verfuegbaren Schulwochen in Reihenfolge.
Private Function ReadWeeks(ByVal st As Worksheet, ByVal startYear As Long, _
                           ByRef nr() As Long, ByRef dVon() As Date, _
                           ByRef dBis() As Date) As Long
    Dim r As Long, n As Long, cap As Long, r1 As Long, r2 As Long
    r1 = modWochenplan.WeekFirstRow()
    r2 = modWochenplan.WeekLastRow()
    cap = r2 - r1 + 1
    If cap < 1 Then cap = 1
    ReDim nr(1 To cap)
    ReDim dVon(1 To cap)
    ReDim dBis(1 To cap)
    n = 0
    For r = r1 To r2
        If Len(Trim$(CStr(st.Cells(r, "A").Value))) > 0 Then
            If IsNumeric(st.Cells(r, "A").Value) Then
                If IsDate(st.Cells(r, "B").Value) And IsDate(st.Cells(r, "C").Value) Then
                    If BoolCell(st.Cells(r, "D").Value) Then
                        n = n + 1
                        nr(n) = CLng(st.Cells(r, "A").Value)
                        dVon(n) = NormDate(CDate(st.Cells(r, "B").Value), startYear)
                        dBis(n) = NormDate(CDate(st.Cells(r, "C").Value), startYear)
                    End If
                End If
            End If
        End If
    Next r
    ReadWeeks = n
End Function


'  Startdatum einer Schulwoche (unabhaengig von "verfuegbar")
Private Function WeekStart(ByVal st As Worksheet, ByVal startYear As Long, _
                           ByVal uw As Long, ByRef dVon As Date) As Boolean
    Dim r As Long
    For r = modWochenplan.WeekFirstRow() To modWochenplan.WeekLastRow()
        If Len(Trim$(CStr(st.Cells(r, "A").Value))) > 0 Then
            If IsNumeric(st.Cells(r, "A").Value) Then
                If CLng(st.Cells(r, "A").Value) = uw Then
                    If IsDate(st.Cells(r, "B").Value) Then
                        dVon = NormDate(CDate(st.Cells(r, "B").Value), startYear)
                        WeekStart = True
                    End If
                    Exit Function
                End If
            End If
        End If
    Next r
End Function


'  Ferieneintraege lesen, aufsteigend nach "von" sortiert
Private Function ReadFerien(ByVal st As Worksheet, ByVal startYear As Long, _
                            ByRef dVon() As Date, ByRef dBis() As Date, _
                            ByRef nam() As String) As Long
    Dim r As Long, n As Long, cap As Long, i As Long, j As Long
    Dim tD As Date, tS As String

    cap = FER_ROWS
    ReDim dVon(1 To cap)
    ReDim dBis(1 To cap)
    ReDim nam(1 To cap)
    n = 0
    For r = modWochenplan.FerFirstRow() To modWochenplan.FerLastRow()
        If IsDate(st.Cells(r, FER_COL_VON).Value) And IsDate(st.Cells(r, FER_COL_BIS).Value) Then
            n = n + 1
            dVon(n) = NormDate(CDate(st.Cells(r, FER_COL_VON).Value), startYear)
            dBis(n) = NormDate(CDate(st.Cells(r, FER_COL_BIS).Value), startYear)
            nam(n) = Trim$(CStr(st.Cells(r, FER_COL_NAME).Value))
        End If
    Next r

    ' einfaches Sortieren nach "von"
    For i = 1 To n - 1
        For j = i + 1 To n
            If dVon(j) < dVon(i) Then
                tD = dVon(i): dVon(i) = dVon(j): dVon(j) = tD
                tD = dBis(i): dBis(i) = dBis(j): dBis(j) = tD
                tS = nam(i): nam(i) = nam(j): nam(j) = tS
            End If
        Next j
    Next i

    ReadFerien = n
End Function


Private Function BoolCell(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        BoolCell = CBool(v)
    Else
        Select Case UCase$(Trim$(CStr(v)))
            Case "WAHR", "TRUE", "JA", "X", "1": BoolCell = True
        End Select
    End If
End Function


'=====================================================================
'  Blatt "Einstellungen" um Fixierungs-Schalter und Ferientabelle
'  erweitern (idempotent, vorhandene Eintraege bleiben erhalten)
'=====================================================================
Public Sub Einstellungen_Erweitern()
    Dim st As Worksheet, fixRw As Long, h As Long

    Set st = modWochenplan.SetSheet()
    If st Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    modWochenplan.FastOn st
    On Error GoTo Fail

    ' --- Fixierungs-Schalter -----------------------------------------
    modWochenplan.SetStep "Fixierungs-Schalter anlegen"
    fixRw = modWochenplan.FixRow()
    If Len(Trim$(CStr(st.Cells(fixRw, "A").Value))) = 0 Then
        st.Cells(fixRw, "A").Value = LBL_FIX
    End If
    ' Schalter als Text "ja"/"nein". Aeltere Mappen haben dort einen
    ' Wahrheitswert stehen - der wird hier einmalig umgestellt.
    With st.Cells(fixRw, "B")
        If Len(Trim$(CStr(.Value))) = 0 Then
            .Value = FIX_NEIN
        ElseIf VarType(.Value) = vbBoolean Then
            .Value = IIf(CBool(.Value), FIX_JA, FIX_NEIN)
        End If
        .HorizontalAlignment = xlCenter
    End With
    On Error Resume Next
    With st.Cells(fixRw, "B").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=FIX_JA & "," & FIX_NEIN
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
    On Error GoTo 0

    ' Der Hinweis "nicht von Hand aendern" kann weg - die Zelle ist
    ' durch den Blattschutz ohnehin gesperrt.
    st.Cells(fixRw, "C").Value = "Wird über die Schaltfläche im Blatt '" & _
                                 CTRL_SHEET & "' gesetzt"

    ' --- Ferientabelle an die richtige Stelle bringen ----------------
    modWochenplan.SetStep "Ferientabelle verschieben"
    MigrateFerienTable st

    modWochenplan.SetStep "Ferientabelle anlegen"
    modWochenplan.RefreshLayout True
    h = modWochenplan.FerHeadRow()

    st.Cells(h, FER_COL_VON).Value = "Ferien von"
    st.Cells(h, FER_COL_BIS).Value = "Ferien bis"
    st.Cells(h, FER_COL_NAME).Value = LBL_FERNAME

    With st.Range(st.Cells(h, FER_COL_VON), st.Cells(h, FER_COL_NAME))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Pattern = xlSolid
        .Interior.Color = modWochenplan.FARBE_HELL
    End With

    With st.Range(st.Cells(h, FER_COL_VON), st.Cells(modWochenplan.FerLastRow(), FER_COL_NAME))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(180, 188, 200)
    End With

    st.Range(st.Cells(modWochenplan.FerFirstRow(), FER_COL_VON), _
             st.Cells(modWochenplan.FerLastRow(), FER_COL_BIS)).NumberFormat = "dd.mm.yyyy"
    st.Columns(FER_COL_VON).ColumnWidth = 12.5
    st.Columns(FER_COL_BIS).ColumnWidth = 12.5
    st.Columns(FER_COL_NAME).ColumnWidth = 24

    ' --- Angaben fuer den Seitenkopf ---------------------------------
    modWochenplan.SetStep "Kopfblock anlegen"
    KopfblockAnlegen st

    ' --- die beiden kleinen Schaltflaechen ---------------------------
    modWochenplan.SetStep "Schaltflaechen in den Einstellungen"
    EnsureEinstellungenButtons st

    modWochenplan.FastOff
    Exit Sub
Fail:
    modWochenplan.ReportError "Erweitern der Einstellungen"
End Sub


'---------------------------------------------------------------------
'  Block "Angaben fuer den Seitenkopf" samt Schulliste anlegen.
'  Vorhandene Eintraege bleiben unangetastet.
'---------------------------------------------------------------------
Private Sub KopfblockAnlegen(ByVal st As Worksheet)
    Dim r0 As Long, kc As Long, sr As Long, sc As Long, adr As String
    Dim i As Long

    ' --- die vier Beschriftungen ------------------------------------
    ' Sie duerfen ueberall stehen, solange sie untereinander in einer
    ' Spalte liegen und der Wert rechts daneben. Fehlen sie, werden sie
    ' unter dem Kopfblock angelegt.
    If modWochenplan.KopfZeileFach() = 0 Then
        r0 = modWochenplan.FixRow() + 3
        If r0 + 3 > modWochenplan.WeekHeadRow() - 1 Then r0 = 1
        If r0 = 1 Then
            kc = KOPF_DEF_COL
        Else
            kc = 1
        End If
        st.Cells(r0, kc).Value = LBL_FACH
        st.Cells(r0 + 1, kc).Value = LBL_KLASSE
        st.Cells(r0 + 2, kc).Value = LBL_SCHULE
        st.Cells(r0 + 3, kc).Value = LBL_LEHRER
        modWochenplan.RefreshLayout True
    End If

    kc = modWochenplan.KopfCol()
    If kc = 0 Then Exit Sub

    ' --- Schulliste ---------------------------------------------------
    If modWochenplan.SchulRow() = 0 Then
        st.Cells(modWochenplan.KopfZeileFach() - 1, KOPF_DEF_COL).Value = LBL_SCHULLISTE
        modWochenplan.RefreshLayout True
    End If
    sr = modWochenplan.SchulRow()
    sc = modWochenplan.SchulCol()
    If sr = 0 Then Exit Sub

    ' Die Beschriftung nannte frueher die Logo-Datei mit. Die Logos
    ' liegen inzwischen als Bilder daneben - der Dateiname verwirrt nur.
    st.Cells(sr, sc).Value = LBL_SCHULLISTE

    If Len(Trim$(CStr(st.Cells(sr + 1, sc).Value))) = 0 Then
        st.Cells(sr + 1, sc).Value = "Schule 1"
        st.Cells(sr + 2, sc).Value = "Schule 2"
        modWochenplan.RefreshLayout True
        sr = modWochenplan.SchulRow()
    End If

    ' --- Aussehen -----------------------------------------------------
    With st.Range(st.Cells(modWochenplan.KopfZeileFach(), kc), _
                  st.Cells(modWochenplan.KopfZeileLehrer(), kc))
        .Font.Bold = False
    End With
    With st.Range(st.Cells(modWochenplan.KopfZeileFach(), kc + 1), _
                  st.Cells(modWochenplan.KopfZeileLehrer(), kc + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(180, 188, 200)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
    End With
    With st.Cells(sr, sc)
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.Color = modWochenplan.FARBE_HELL
    End With
    st.Columns(sc + 1).ColumnWidth = 26

    ' --- Auswahlliste fuer die Schule ---------------------------------
    If modWochenplan.SchulLastRow() >= sr + 1 Then
        adr = "=" & SET_SHEET & "!" & _
              st.Range(st.Cells(sr + 1, sc), _
                       st.Cells(modWochenplan.SchulLastRow(), sc)).Address(True, True)
        On Error Resume Next
        With st.Cells(modWochenplan.KopfZeileSchule(), kc + 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=adr
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = "Unbekannte Schule"
            .ErrorMessage = "Bitte eine Schule aus der Liste wählen."
        End With
        On Error GoTo 0
    End If

    ' --- Reste der frueheren Logo-Verwaltung aufraeumen ----------------
    ' Die Logos liegen inzwischen fest in der Mappe (Formen
    ' wpKopfLogo_<n> in "Wochenplan" und "Lernbereiche"). Der ganze
    ' Apparat, der sie fruener aus dem Ordner "Bilder" gesucht und
    ' eingelesen hat, ist ersatzlos entfallen.
    DateinamenSpalteLeeren st

    ' --- interne Arbeitsdaten ausblenden ------------------------------
    ArbeitsbereichVerstecken st
End Sub


'---------------------------------------------------------------------
'  Stundenliste, Schulliste und die eingebetteten Logos sind reine
'  Arbeitsdaten. Der Nutzer muss sie nicht sehen - er waehlt Stunde
'  und Schule ueber die Auswahllisten.
'
'  Ausgeblendet wird, NICHT verschoben: an der Stundenliste haengen
'  die Regeln der bedingten Formatierung im Wochenplan und die
'  Gueltigkeitsliste der Spalte K. Beim Verschieben per VBA ziehen
'  diese Bezuege nicht mit - beim Ausblenden aendert sich gar nichts.
'  Gueltigkeitslisten und Formeln lesen aus ausgeblendeten Spalten
'  voellig normal weiter.
'---------------------------------------------------------------------
Private Sub ArbeitsbereichVerstecken(ByVal st As Worksheet)
    Dim c1 As Long, c2 As Long, c As Long

    If st Is Nothing Then Exit Sub

    c1 = 0: c2 = 0
    Merke modWochenplan.StdCol(), c1, c2
    Merke modWochenplan.StdCol() - 1, c1, c2      ' die Nummernspalte davor
    Merke modWochenplan.SchulCol(), c1, c2
    Merke modWochenplan.SchulCol() + 1, c1, c2    ' Spalte mit den Logos
    If c1 = 0 Then Exit Sub

    ' Nur ausblenden, wenn der Bereich rechts von den sichtbaren
    ' Daten liegt - sonst verschwaende eine Fehlerkennung Spalten,
    ' die der Nutzer braucht.
    If c1 <= st.Cells(1, FER_COL_NAME).Column Then Exit Sub

    On Error Resume Next
    For c = c1 To c2
        st.Columns(c).Hidden = True
    Next c
    On Error GoTo 0
End Sub


'  Spaltennummer in die Spanne c1..c2 aufnehmen.
Private Sub Merke(ByVal c As Long, ByRef c1 As Long, ByRef c2 As Long)
    If c <= 0 Then Exit Sub
    If c1 = 0 Then
        c1 = c: c2 = c
        Exit Sub
    End If
    If c < c1 Then c1 = c
    If c > c2 Then c2 = c
End Sub


'---------------------------------------------------------------------
'  Die alte Spalte mit den Logo-Dateinamen leeren. Die Logos stehen
'  jetzt als Bild daneben; der Dateiname verwirrt nur.
'---------------------------------------------------------------------
Private Sub DateinamenSpalteLeeren(ByVal st As Worksheet)
    Dim i As Long, sr As Long, sc As Long

    sr = modWochenplan.SchulRow()
    sc = modWochenplan.SchulCol()
    If sr = 0 Or sc = 0 Then Exit Sub

    For i = sr + 1 To modWochenplan.SchulLastRow()
        If Not IsEmpty(st.Cells(i, sc + 1).Value) Then
            If InStr(1, CStr(st.Cells(i, sc + 1).Value), ".", vbTextCompare) > 0 Then
                st.Cells(i, sc + 1).ClearContents
            End If
        End If
    Next i
End Sub


'---------------------------------------------------------------------
'  Zwei kleine Schaltflaechen im Blatt "Einstellungen":
'  eine ueber der Ferientabelle, eine unter dem Kopfblock.
'---------------------------------------------------------------------
Private Sub EnsureEinstellungenButtons(ByVal st As Worksheet)
    Dim h As Long, kr As Long, kc As Long, sr As Long, sc As Long

    h = modWochenplan.FerHeadRow()
    If h > 1 Then
        st.Rows(h - 1).RowHeight = 22
        SmallButton st, "esBtnFerien", _
                    st.Cells(h - 1, FER_COL_VON), st.Cells(h - 1, FER_COL_NAME), _
                    "Ferien aus Kalender vorschlagen", _
                    "modKalender.Ferien_Aus_Kalender_Vorschlagen", _
                    RGB(219, 237, 222), RGB(93, 143, 102)
    End If

    kr = modWochenplan.KopfZeileLehrer()
    kc = modWochenplan.KopfCol()
    If kr > 0 And kc > 0 Then
        st.Rows(kr + 1).RowHeight = 22
        SmallButton st, "esBtnKopf", _
                    st.Cells(kr + 1, kc), st.Cells(kr + 1, kc + 1), _
                    "Kopf übernehmen", "modKopf.Kopf_Aktualisieren", _
                    RGB(219, 237, 222), RGB(93, 143, 102)
    End If

    ' "Logo waehlen" hat hier keine Schaltflaeche mehr: die Logos
    ' stecken versteckt in der Mappe und werden ueber die Auswahl bei
    ' "Schule" genommen. Fuer eine zusaetzliche Schule gibt es das
    ' Makro weiterhin ueber Alt+F8 (Logo_Waehlen).
    LoescheForm st, "esBtnLogo"
End Sub


'  Eine Schaltflaeche entfernen, die es nicht mehr geben soll.
Private Sub LoescheForm(ByVal ws As Worksheet, ByVal nm As String)
    Dim i As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = nm Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub


Private Sub SmallButton(ByVal ws As Worksheet, ByVal nm As String, _
                        ByVal vonZelle As Range, ByVal bisZelle As Range, _
                        ByVal caption As String, ByVal macro As String, _
                        ByVal fillCol As Long, ByVal lineCol As Long)
    Dim s As Shape, i As Long

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = nm Then ws.Shapes(i).Delete
    Next i

    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                vonZelle.Left + 1, vonZelle.Top + 1, _
                (bisZelle.Left + bisZelle.Width) - vonZelle.Left - 2, _
                vonZelle.Height - 2)
    If s Is Nothing Then Exit Sub
    s.Name = nm
    With s
        .Placement = xlMoveAndSize
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = fillCol
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = lineCol
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
        .OnAction = macro
        .AlternativeText = caption
    End With
    With s.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(45, 55, 72)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .MarginLeft = 2
        .MarginRight = 2
        .AutoSize = msoAutoSizeNone
    End With
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------
'  Sucht eine bereits vorhandene Ferientabelle irgendwo im Blatt und
'  verschiebt sie an die vorgesehene Stelle (Spalten I/J/K, Kopfzeile
'  auf Hoehe der Kalender-Kopfzeile). Vorhandene Eintraege bleiben
'  erhalten. Ist die Tabelle schon am richtigen Platz, passiert nichts.
'---------------------------------------------------------------------
Private Sub MigrateFerienTable(ByVal st As Worksheet)
    Dim r As Long, c As Long, foundR As Long, foundC As Long
    Dim tgtR As Long, tgtC As Long, i As Long
    Dim vVon() As Variant, vBis() As Variant, vNam() As Variant
    Dim v As String

    tgtR = modWochenplan.FerHeadRow()
    tgtC = st.Cells(1, FER_COL_VON).Column

    foundR = 0: foundC = 0
    For c = 4 To 30
        For r = 1 To 120
            v = LCase$(Trim$(CStr(st.Cells(r, c).Value)))
            If v = LCase$(LBL_FERVON) Then
                foundR = r
                foundC = c
                Exit For
            End If
        Next r
        If foundR > 0 Then Exit For
    Next c

    If foundR = 0 Then Exit Sub                       ' noch keine Tabelle da
    If foundR = tgtR And foundC = tgtC Then Exit Sub  ' schon am richtigen Platz

    ReDim vVon(1 To FER_ROWS)
    ReDim vBis(1 To FER_ROWS)
    ReDim vNam(1 To FER_ROWS)
    For i = 1 To FER_ROWS
        vVon(i) = st.Cells(foundR + i, foundC).Value
        vBis(i) = st.Cells(foundR + i, foundC + 1).Value
        vNam(i) = st.Cells(foundR + i, foundC + 2).Value
    Next i

    st.Range(st.Cells(foundR, foundC), st.Cells(foundR + FER_ROWS, foundC + 2)).Clear
    For i = 0 To 2
        st.Columns(foundC + i).ColumnWidth = st.StandardWidth
    Next i

    For i = 1 To FER_ROWS
        st.Cells(tgtR + i, tgtC).Value = vVon(i)
        st.Cells(tgtR + i, tgtC + 1).Value = vBis(i)
        st.Cells(tgtR + i, tgtC + 2).Value = vNam(i)
    Next i
End Sub


'=====================================================================
'  Ferientabelle aus den Luecken der Schulwochen-Tabelle vorbefuellen
'=====================================================================
Public Sub Ferien_Aus_Kalender_Vorschlagen()
    Dim st As Worksheet, startYear As Long
    Dim r As Long, tgt As Long, n As Long
    Dim d1 As Date, d2 As Date
    Dim vorhanden As Boolean

    Set st = modWochenplan.SetSheet()
    If st Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    For r = modWochenplan.FerFirstRow() To modWochenplan.FerLastRow()
        If IsDate(st.Cells(r, FER_COL_VON).Value) Then vorhanden = True
    Next r
    If vorhanden Then
        If modWochenplan.Frage("In der Ferientabelle stehen bereits Einträge." & vbCrLf & vbCrLf & _
                  "Sollen sie durch einen neuen Vorschlag aus dem Schulwochen-Kalender " & _
                  "ersetzt werden?", vbQuestion + vbOKCancel + vbDefaultButton2, _
                  "Stoffverteilungsplan") <> vbOK Then Exit Sub
    End If

    modWochenplan.FastOn st
    On Error GoTo Fail

    startYear = SchoolStartYear(st)
    st.Range(st.Cells(modWochenplan.FerFirstRow(), FER_COL_VON), _
             st.Cells(modWochenplan.FerLastRow(), FER_COL_NAME)).ClearContents

    tgt = modWochenplan.FerFirstRow()
    For r = modWochenplan.WeekFirstRow() To modWochenplan.WeekLastRow() - 1
        If IsDate(st.Cells(r, "C").Value) And IsDate(st.Cells(r + 1, "B").Value) Then
            d1 = NormDate(CDate(st.Cells(r, "C").Value), startYear)
            d2 = NormDate(CDate(st.Cells(r + 1, "B").Value), startYear)
            If d2 - d1 > 1 Then
                If tgt <= modWochenplan.FerLastRow() Then
                    st.Cells(tgt, FER_COL_VON).Value = CDate(d1 + 1)
                    st.Cells(tgt, FER_COL_BIS).Value = CDate(d2 - 1)
                    st.Cells(tgt, FER_COL_NAME).Value = FerienName(CDate(d1 + 1))
                    tgt = tgt + 1
                    n = n + 1
                End If
            End If
        End If
    Next r

    modWochenplan.FastOff
    ' Diese Schaltflaeche sitzt im Blatt "Einstellungen" - dort gehoert
    ' der Anwender danach auch wieder hin.
    modWochenplan.GotoSheet SET_SHEET
    modWochenplan.Info n & " Zeitraum/Zeiträume übernommen - bitte die " & _
           "Bezeichnungen prüfen."
    Exit Sub
Fail:
    modWochenplan.ReportError "Vorschlagen der Ferien"
End Sub


Private Function FerienName(ByVal d As Date) As String
    Select Case Month(d)
        Case 10, 11:  FerienName = "Herbstferien"
        Case 12, 1:   FerienName = "Weihnachtsferien"
        Case 2:       FerienName = "Faschingsferien"
        Case 3, 4:    FerienName = "Osterferien"
        Case 5, 6:    FerienName = "Pfingstferien"
        Case 7, 8, 9: FerienName = "Sommerferien"
        Case Else:    FerienName = "unterrichtsfrei"
    End Select
End Function
