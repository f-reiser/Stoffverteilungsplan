Attribute VB_Name = "modSteuerung"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Blatt "Steuerung"
'  ------------------------------------------------------------------
'  Auf dem Blatt stehen nur die Schaltflaechen, die man wirklich
'  braucht:
'    oben    was man im Alltag benutzt
'    unten   was man ein- oder zweimal im Jahr benutzt
'
'  Bewusst KEINE Schaltflaechen mehr fuer:
'    Wochenplan aktualisieren  - rufen die Makros selbst auf, wenn es
'                                noetig ist; von Hand kann durch den
'                                Blattschutz ohnehin nichts verrutschen
'    Ferien vorschlagen        - sitzt jetzt direkt ueber der
'                                Ferientabelle im Blatt "Einstellungen"
'    Kopf uebernehmen          - sitzt beim Kopfblock in den
'                                "Einstellungen"
'    Anleitungsblaetter        - wuerde von Hand gemachte Aenderungen
'                                an der Anleitung ueberschreiben; nur
'                                ueber Alt+F8, Eintrag
'                                "Anleitungsblatt_neu_schreiben"
'    Diagnose                  - nur im Fehlerfall, ueber Alt+F8
'                                ("Diagnose" in modStart)
'=====================================================================

Public Const CTRL_SHEET As String = "Steuerung"

' Zellen mit den beiden Statuszeilen
Public Const CTRL_STATUS_CODE   As String = "B4"
Public Const CTRL_STATUS_SCHUTZ As String = "B5"

'  Beschriftungen der beiden Fixier-Schaltflaechen. Sie dienen
'  gleichzeitig als Wiedererkennung (AlternativeText der Form), wenn
'  die gerade nicht sinnvolle von beiden ausgegraut wird.
Public Const CAP_FIXIEREN As String = "Stoffverteilungsplan fixieren"
Public Const CAP_AUFHEBEN As String = "Fixierung aufheben"
Public Const CAP_PDF      As String = "Als PDF exportieren"
Public Const CAP_IMPORT   As String = "Daten importieren"
Public Const CAP_SETUP    As String = "Einrichtung / Reparatur"

Private Const BTN_PREFIX As String = "stBtn_"
Private Const BTN_COL    As String = "B"
Private Const TXT_COL    As String = "C"


'=====================================================================
'  Der Knopf "Einrichtung / Reparatur".
'
'  Er ruft nur noch nach: seit es "Daten importieren" gibt, ist er
'  fuer ein Update nicht mehr der Weg, und wer ihn aus Gewohnheit
'  drueckt, soll das einmal lesen. Setup_Stoffverteilungsplan selbst
'  bleibt ohne Rueckfrage - modUebernahme ruft es am Ende auf, und
'  dort waere eine zweite Frage nur laestig.
'=====================================================================
Public Sub Einrichtung_Reparatur()
    If modWochenplan.Frage( _
           "Für ein Update auf eine neue Fassung ist das nicht der vorgesehene Weg." & _
           vbCrLf & vbCrLf & _
           "Gedacht ist: die neue, leere Datei öffnen und dort """ & CAP_IMPORT & _
           """ drücken. Dieser Knopf hier baut nur die Formeln und den Blattschutz " & _
           "dieser Mappe neu auf - er holt keine neuen Makros." & vbCrLf & vbCrLf & _
           "Trotzdem neu einrichten?", _
           vbOKCancel) <> vbOK Then Exit Sub
    Setup_Stoffverteilungsplan
End Sub


'=====================================================================
'  Komplette Einrichtung
'=====================================================================
Public Sub Setup_Stoffverteilungsplan()
    Dim ws As Worksheet

    On Error GoTo Fail
    If Not modWochenplan.LayoutReady() Then Exit Sub

    modKalender.Einstellungen_Erweitern
    EnsureSteuerung
    modAnleitung.Anleitungsblaetter_Sicherstellen

    '  Bewusst EIN FastOn ueber beide Blaetter: FastOn/FastOff sind
    '  nicht schachtelbar, und Lernbereiche_Aufbauen schreibt in ein
    '  geschuetztes Blatt.
    modWochenplan.FastOn
    Set ws = modWochenplan.WpSheet()
    If Not ws Is Nothing Then
        modWochenplan.SetStep "Wochenplan aufbauen"
        modWochenplan.RebuildAll ws, modWochenplan.PlanLastRow(ws)
        modWochenplan.SetStep "Zeilen-Schaltflaechen anlegen"
        modWochenplan.EnsureRowButtons
    End If
    modWochenplan.SetStep "Lernbereiche berechnen"
    modWochenplan.Lernbereiche_Aufbauen
    modWochenplan.FastOff

    modKopf.Kopf_Aktualisieren True
    modSchutz.Blattschutz_Einrichten

    '  Zum Schluss und nicht am Anfang: nach dem Neuaufbau steht in der
    '  Warnzeile, was der Neuaufbau NICHT heilen konnte.
    modPruefung.WarnungAnzeigen

    modWochenplan.GotoSheet CTRL_SHEET, "B2"

    modWochenplan.Info "Die Einrichtung ist abgeschlossen."
    Exit Sub
Fail:
    modWochenplan.ReportError "Einrichten"
End Sub


'=====================================================================
'  Blatt "Steuerung" anlegen bzw. auffrischen
'=====================================================================
Private Sub EnsureSteuerung()
    Dim ws As Worksheet, r As Long

    modWochenplan.FastOn
    On Error GoTo Fail

    modWochenplan.SetStep "Blatt Steuerung anlegen"
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CTRL_SHEET)
    On Error GoTo Fail
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = CTRL_SHEET
    End If

    modWochenplan.SetStep "Steuerung: alte Schaltflaechen entfernen"
    RemoveOldButtons ws

    modWochenplan.SetStep "Steuerung: Blattaufbau"
    ws.Cells.Clear
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Fail

    ws.Columns("A").ColumnWidth = 3
    ws.Columns(BTN_COL).ColumnWidth = 32
    ws.Columns(TXT_COL).ColumnWidth = 86
    ws.Rows(1).RowHeight = 14

    With ws.Range("B2")
        .Value = "Stoffverteilungsplan"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(45, 55, 72)
    End With
    ws.Rows(2).RowHeight = 30

    With ws.Range("B3")
        .Value = "Die vier kleinen Schaltflächen zum Verschieben, Einfügen und " & _
                 "Löschen von Zeilen erscheinen direkt im Wochenplan, sobald dort " & _
                 "eine Zeile markiert ist."
        .Font.Color = modWochenplan.FARBE_LEISE
    End With

    With ws.Range(CTRL_STATUS_CODE)
        ' LOWER(): der Schalter steht als Text "ja"/"nein" da. Aeltere
        ' Mappen mit WAHR/FALSCH liefern hier "falsch" bzw. "wahr" und
        ' landen damit richtig im Zweig "offen".
        .Formula = "=""Referenz-Codes: ""&IF(LOWER(" & SET_SHEET & "!" & _
                   modWochenplan.FixCellAddr() & ")=""" & FIX_JA & _
                   """,""fixiert"",""offen - werden noch berechnet"")"
        .Font.Bold = True
        .Font.Color = modWochenplan.FARBE_LEISE
    End With
    With ws.Range(CTRL_STATUS_SCHUTZ)
        .Value = "Blattschutz: " & IIf(modSchutz.SchutzAktiv(), "aktiv", "AUS")
        .Font.Bold = True
        .Font.Color = modWochenplan.FARBE_LEISE
    End With

    ' ---------------- Alltag -----------------------------------------
    r = 7
    Ueberschrift ws, r, "Im Alltag"
    r = r + 1

    AddButton ws, r, True, "Wochenplan neu aufbauen", _
              "modKalender.UW_Und_Ferien_Generieren", _
              RGB(219, 237, 222), RGB(93, 143, 102), _
              "Der Knopf, mit dem die Planung beginnt - und mit dem man sie jederzeit " & _
              "wieder gerade rückt. Er trägt im Blatt """ & WP_SHEET & """ in Spalte E " & _
              "(""UW"") die verfügbaren Unterrichtswochen der Reihe nach ein, setzt die " & _
              "grauen Ferienzeilen an die passenden Stellen und füllt Datum (C) und " & _
              "KW (D) neu. BESTANDSDATEN BLEIBEN ERHALTEN: was in den Spalten F bis M " & _
              "steht - Lehrplan-Code, Thema, Kompetenzen, Material, Erledigt-Haken, " & _
              "Stunde, Hinweis, Notizen - wandert mit seiner Zeile mit und wird nicht " & _
              "angetastet. Es ändert sich nur, welche Unterrichtswoche neben welcher " & _
              "Zeile steht. Nach jedem Einfügen, Löschen oder Verschieben von Zeilen " & _
              "und nach jeder Änderung an der Ferientabelle einmal drücken."
    r = r + 2

    AddButton ws, r, True, CAP_FIXIEREN, _
              "modSteuerung.Plan_Fixieren", _
              RGB(252, 240, 214), RGB(191, 148, 44), _
              "Friert im Blatt """ & WP_SHEET & """ die Referenz-Codes in Spalte B ein: " & _
              "aus den Formeln werden feste Texte. Ab dann wandert der Code beim " & _
              "Verschieben mit seiner Zeile mit und ändert sich auch nicht mehr, wenn " & _
              "oberhalb Zeilen dazukommen oder wegfallen. Drücken, sobald die Planung " & _
              "steht und du die Codes weitergibst - etwa auf Arbeitsblättern."
    r = r + 2

    AddButton ws, r, True, CAP_AUFHEBEN, _
              "modSteuerung.Fixierung_Aufheben", _
              RGB(247, 223, 223), RGB(168, 96, 96), _
              "Macht die Referenz-Codes in Spalte B des Blattes """ & WP_SHEET & """ " & _
              "wieder zu Formeln. ACHTUNG: dabei werden ALLE Codes neu durchnummeriert. " & _
              "Bereits weitergegebene Codes stimmen danach nicht mehr - nur benutzen, " & _
              "solange niemand mit ihnen arbeitet."
    r = r + 2

    AddButton ws, r, True, CAP_PDF, _
              "modKopf.PDF_Export", _
              RGB(219, 229, 245), RGB(84, 118, 176), _
              "Schreibt beide Blätter in EINE PDF-Datei: zuerst """ & WP_SHEET & """ " & _
              "(Spalten B bis M), dahinter """ & LB_SHEET & """ (Spalten A bis J). " & _
              "Immer quer, die Breite wird auf eine Seite skaliert. Ordner und " & _
              "Dateiname fragt der gewohnte Speichern-Dialog ab; den zuletzt " & _
              "benutzten Ordner merkt sich die Datei."
    r = r + 2

    PdfFormatZelle ws, r
    r = r + 2

    ' ---------------- Update ------------------------------------------
    r = r + 1
    Ueberschrift ws, r, "Update"
    r = r + 1

    AddButton ws, r, True, CAP_IMPORT, _
              "modUebernahme.Daten_Uebernehmen", _
              RGB(224, 234, 240), RGB(90, 132, 158), _
              "DER VORGESEHENE WEG FÜR EIN UPDATE. Eine neue Fassung der Mappe " & _
              "kommt als fertige, leere Datei und bringt die aktuellen Makros schon " & _
              "mit. Dieser Knopf holt deine bisherigen Inhalte aus der alten Datei " & _
              "herein und baut den Plan fertig auf; die alte Datei wird dabei nur " & _
              "gelesen. Danach nur noch speichern. Geht nur in eine Mappe, in der " & _
              "noch nichts geplant ist. Was im Einzelnen übernommen wird, steht im " & _
              "Blatt """ & modAnleitung.SHEET_HELP & """ unter ""Eine neue Fassung " & _
              "der Mappe""."
    r = r + 2

    ' ---------------- Selten -----------------------------------------
    r = r + 1
    Ueberschrift ws, r, "Selten gebraucht"
    r = r + 1

    AddButton ws, r, False, CAP_SETUP, _
              "modSteuerung.Einrichtung_Reparatur", _
              RGB(232, 236, 242), RGB(150, 160, 178), _
              "Für ein Update NICHT nötig - dafür ist """ & CAP_IMPORT & """ da. " & _
              "Baut dieses Blatt, die Hilfsangaben im Blatt """ & SET_SHEET & """, die " & _
              "Formelspalten des """ & WP_SHEET & """ (A bis D, L und die " & _
              "ausgeblendeten Hilfsspalten rechts) sowie die Rechenspalten F bis I und " & _
              "die Summenzeile in """ & LB_SHEET & """ neu auf und setzt den " & _
              "Blattschutz frisch. Für den Spezialfall, dass etwas verrutscht ist oder " & _
              "eine Fehlermeldung darum bittet."
    r = r + 2

    PruefHinweisZelle ws, r
    r = r + 2

    '  Bewusst rot und mit Warnzeichen: das ist die einzige
    '  Schaltflaeche der Mappe, mit der man sich ernsthaft schaden
    '  kann. Siehe die Begruendung im Text.
    AddButton ws, r, False, ChrW$(9888) & " Blattschutz ein / aus", _
              "modSchutz.Blattschutz_Umschalten", _
              RGB(250, 226, 226), RGB(186, 74, 74), _
              "Hebt den Schutz aller Blätter auf. Danach ist alles änderbar - auch " & _
              "das, was die Makros brauchen. Zwei Dinge gehen dabei erfahrungsgemäß " & _
              "schief: Erstens sind eigene Umbauten beim nächsten Update WEG, denn " & _
              "übernommen werden nur deine Inhalte, nicht dein Blattaufbau. Zweitens " & _
              "reicht ein Sortieren, ein Ausschneiden-und-Einfügen oder eine " & _
              "gelöschte Spalte, um die Formelbezüge reihum zu zerlegen - danach " & _
              "rechnet nichts mehr richtig, und man sieht es erst Wochen später. " & _
              "Genau davor schützt der Blattschutz. Nur einschalten, wenn du weißt, " & _
              "was du tust - und danach sofort wieder zurückstellen.", _
              "ACHTUNG:"
    r = r + 2

    r = r + 1
    '  Ueber beide Spalten, sonst steht der Satz in der 32 Zeichen
    '  schmalen Schaltflaechenspalte und ist abgeschnitten.
    With ws.Range(ws.Cells(r, BTN_COL), ws.Cells(r, TXT_COL))
        .Merge
        .Value = "Zwei weitere Schaltflächen sitzen dort, wo sie gebraucht werden, " & _
                 "beide im Blatt """ & SET_SHEET & """: ""Ferien aus Kalender " & _
                 "vorschlagen"" direkt über der Ferientabelle (Spalten I bis K) und " & _
                 """Kopf übernehmen"" bei den vier Kopfangaben Fach, Klasse, Schule " & _
                 "und Lehrkraft."
        .WrapText = True
        .VerticalAlignment = xlTop
        .IndentLevel = 1
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(140, 150, 168)
    End With
    ws.Rows(r).RowHeight = 28

    ws.Range("B2").Select
    modWochenplan.FastOff
    KnoepfeAktualisieren
    Exit Sub
Fail:
    modWochenplan.ReportError "Aufbauen des Blattes " & CTRL_SHEET
End Sub


Private Sub Ueberschrift(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    ws.Rows(r).RowHeight = 24
    With ws.Cells(r, BTN_COL)
        .Value = txt
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = modWochenplan.FARBE_BALKEN
    End With
    With ws.Range(ws.Cells(r, BTN_COL), ws.Cells(r, TXT_COL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 208, 220)
    End With
End Sub


Private Sub RemoveOldButtons(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub


'  gross = True  -> zwei Zeilen hoch, groessere Schrift (Alltag)
'  gross = False -> eine Zeile hoch (selten gebraucht)
'---------------------------------------------------------------------
'  Auswahlliste fuer das Papierformat, direkt unter dem PDF-Knopf.
'  Damit braucht der Export keine Rueckfrage per Meldungsfenster.
'  Die Zelle bekommt den Namen wpPdfFormat - so findet modKopf sie
'  wieder, ohne dass irgendwo eine Zeilennummer steht.
'---------------------------------------------------------------------
Private Sub PdfFormatZelle(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Range

    ws.Rows(r).RowHeight = 20
    Set c = ws.Cells(r, BTN_COL)

    With c
        .Value = "DIN A4"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = modWochenplan.FARBE_KOPF
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
        .Locked = False
    End With
    With c.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = modWochenplan.FARBE_KOPF_HELL
    End With

    On Error Resume Next
    With c.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="DIN A4,DIN A3"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
    ThisWorkbook.Names.Add Name:=modKopf.PDF_FORMAT_NAME, _
                           RefersTo:="=" & ws.Name & "!" & c.Address(True, True)
    On Error GoTo 0

    With ws.Cells(r, TXT_COL)
        .Value = "Papierformat für den Export. DIN A3 lohnt sich, wenn die Spalten " & _
                 "auf A4 zu schmal zusammengedrückt werden."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = modWochenplan.FARBE_LEISE
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
End Sub


'---------------------------------------------------------------------
'  Die Zeile, in der modPruefung seine Warnung ablegt - direkt unter
'  "Einrichtung / Reparatur", weil genau dieser Knopf die Antwort auf
'  eine Warnung ist.
'
'  Sie bekommt den Namen wpPruefHinweis: so findet modPruefung sie
'  wieder, ohne dass irgendwo eine Zeilennummer steht. Feste Hoehe,
'  keine Anpassung an den Text - verbundene Zellen koennen kein
'  AutoFit, und die Hoehe liesse sich bei geschuetztem Blatt ohnehin
'  nicht nachziehen.
'---------------------------------------------------------------------
Private Sub PruefHinweisZelle(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Range

    ws.Rows(r).RowHeight = 30
    ws.Rows(r + 1).RowHeight = 10

    Set c = ws.Cells(r, BTN_COL)
    With ws.Range(c, ws.Cells(r, TXT_COL))
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .IndentLevel = 1
        .Font.Size = 9
    End With
    c.Value = modPruefung.PRF_SAUBER
    c.Font.Color = modWochenplan.FARBE_LEISE

    On Error Resume Next
    ThisWorkbook.Names.Add Name:=modPruefung.PRF_ZELLE_NAME, _
                           RefersTo:="=" & ws.Name & "!" & c.Address(True, True)
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------
'  Von den beiden Fixier-Schaltflaechen ist immer nur eine sinnvoll.
'  Die andere wird ausgegraut - sie bleibt anklickbar, sagt dann aber
'  in einem Satz, dass es nichts zu tun gibt.
'---------------------------------------------------------------------
Private Sub KnoepfeAktualisieren()
    Dim ws As Worksheet, fixiert As Boolean

    ' SheetOrNothing gibt es nur als Private in modAnleitung und
    ' modSchutz - von hier aus also nicht erreichbar.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CTRL_SHEET)
    If ws Is Nothing Then Exit Sub

    fixiert = modWochenplan.IsPlanFixed()
    KnopfFarbe ws, CAP_FIXIEREN, Not fixiert, _
               RGB(252, 240, 214), RGB(191, 148, 44)
    KnopfFarbe ws, CAP_AUFHEBEN, fixiert, _
               RGB(247, 223, 223), RGB(168, 96, 96)
    On Error GoTo 0
End Sub


Private Sub KnopfFarbe(ByVal ws As Worksheet, ByVal caption As String, _
                       ByVal aktiv As Boolean, ByVal fillCol As Long, _
                       ByVal lineCol As Long)
    Dim i As Long, s As Shape

    On Error Resume Next
    For i = 1 To ws.Shapes.Count
        Set s = ws.Shapes(i)
        If s.AlternativeText = caption Then
            If aktiv Then
                s.Fill.ForeColor.RGB = fillCol
                s.Line.ForeColor.RGB = lineCol
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = modWochenplan.FARBE_TEXT
            Else
                s.Fill.ForeColor.RGB = RGB(240, 242, 245)
                s.Line.ForeColor.RGB = RGB(205, 210, 218)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(168, 176, 188)
            End If
            Exit For
        End If
    Next i
End Sub


'---------------------------------------------------------------------
'  Eine Schaltflaeche samt Erklaerungstext daneben.
'
'  gross = True  -> hoehere Schaltflaeche, groessere Schrift (Alltag)
'  gross = False -> flache Schaltflaeche (selten gebraucht)
'
'  Jede Schaltflaeche belegt GENAU EINE Zeile, darunter kommt eine
'  schmale Abstandszeile. Die Hoehe dieser einen Zeile richtet sich
'  nach dem Erklaerungstext, nicht umgekehrt: die Texte sind teils
'  mehrere Saetze lang, und verbundene Zellen koennen kein AutoFit -
'  eine feste Hoehe hat den Text abgeschnitten. Die Schaltflaeche
'  selbst behaelt ihre Hoehe und haengt oben in der Zeile.
'---------------------------------------------------------------------
Private Sub AddButton(ByVal ws As Worksheet, ByVal r As Long, ByVal gross As Boolean, _
                      ByVal caption As String, ByVal macro As String, _
                      ByVal fillCol As Long, ByVal lineCol As Long, _
                      ByVal hint As String, Optional ByVal warnung As String = "")
    Dim s As Shape, cel As Range, nm As String
    Dim btnHoehe As Double, txtGroesse As Double, hoehe As Double

    If Len(warnung) > 0 Then hint = warnung & "  " & hint

    If gross Then
        btnHoehe = 46
        txtGroesse = 10
    Else
        btnHoehe = 26
        txtGroesse = 9
    End If

    hoehe = TextHoehe(ws, hint, txtGroesse)
    If hoehe < btnHoehe + 6 Then hoehe = btnHoehe + 6
    ws.Rows(r).RowHeight = hoehe
    ws.Rows(r + 1).RowHeight = 10

    Set cel = ws.Cells(r, BTN_COL)
    nm = BTN_PREFIX & r

    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                               cel.Left + 2, cel.Top + 2, _
                               ws.Columns(BTN_COL).Width - 6, btnHoehe)
    s.Name = nm
    With s
        '  xlMove und nicht xlMoveAndSize: die Zeile ist absichtlich
        '  hoeher als die Schaltflaeche - mit xlMoveAndSize wuerde
        '  Excel sie beim naechsten Zeilenhoehen-Wechsel mitziehen.
        .Placement = xlMove
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = fillCol
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = lineCol
        .Line.Weight = IIf(gross, 1.25, 0.75)
        .Shadow.Visible = msoFalse
        .OnAction = macro
        .AlternativeText = caption
    End With
    With s.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = IIf(gross, 11.5, 10)
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(45, 55, 72)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .MarginLeft = 4
        .MarginRight = 4
        .AutoSize = msoAutoSizeNone
    End With

    With ws.Cells(r, TXT_COL)
        .Value = hint
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Font.Size = txtGroesse
        .Font.Color = RGB(100, 110, 128)
        '  Characters(...) faerbt nur den Anfang. Das geht ausdruecklich
        '  NICHT bei verbundenen Zellen - deshalb ist die Textspalte
        '  seit dem Umbau eine einzelne Zelle.
        If Len(warnung) > 0 Then
            On Error Resume Next
            .Characters(1, Len(warnung)).Font.Color = RGB(186, 74, 74)
            .Characters(1, Len(warnung)).Font.Bold = True
            On Error GoTo 0
        End If
    End With
End Sub


'  Geschaetzte Hoehe in Punkt, die der Text in der Erklaerungsspalte
'  braucht. Die Spaltenbreite ist in Zeichen der Standardschrift
'  (11 pt) angegeben; bei kleinerer Schrift passen entsprechend mehr
'  Zeichen hinein. Der Abschlag faengt breite Zeichen und den
'  Zeilenumbruch an Wortgrenzen ab - lieber eine Zeile zu viel als
'  ein abgeschnittener Satz.
Private Function TextHoehe(ByVal ws As Worksheet, ByVal txt As String, _
                           ByVal groesse As Double) As Double
    Dim proZeile As Double, zeilen As Long

    proZeile = ws.Columns(TXT_COL).ColumnWidth * (11# / groesse) * 0.88
    If proZeile < 10 Then proZeile = 10
    zeilen = Int(Len(txt) / proZeile) + 1
    TextHoehe = zeilen * groesse * 1.35 + 8
End Function


'=====================================================================
'  Fixieren / Fixierung aufheben
'=====================================================================
Public Sub Plan_Fixieren()
    Dim ws As Worksheet, st As Worksheet, lastRow As Long, r As Long, n As Long
    Dim fixRw As Long

    Set ws = modWochenplan.WpSheet()
    Set st = modWochenplan.SetSheet()
    If ws Is Nothing Or st Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    If modWochenplan.IsPlanFixed() Then
        modWochenplan.Info "Der Plan ist bereits fixiert."
        modWochenplan.GotoSheet CTRL_SHEET
        Exit Sub
    End If

    If modWochenplan.Frage("Die Referenz-Codes in Spalte B werden zu festen Werten. " & _
              "Fortfahren?", vbQuestion + vbOKCancel) <> vbOK Then Exit Sub

    modWochenplan.FastOn
    On Error GoTo Fail

    modWochenplan.SetStep "Referenz-Codes einfrieren"
    lastRow = modWochenplan.PlanLastRow(ws)
    Application.Calculate
    For r = WP_FIRST_ROW To lastRow
        If Not modWochenplan.IsFerienRow(ws, r) Then
            ws.Cells(r, "B").Value = ws.Cells(r, "B").Value
            If Len(Trim$(CStr(ws.Cells(r, "B").Value))) > 0 Then n = n + 1
        End If
    Next r

    fixRw = modWochenplan.FixRow()
    st.Cells(fixRw, "B").Value = FIX_JA
    st.Cells(fixRw, "C").Value = "fixiert am " & Format$(Date, "dd.mm.yyyy")

    modWochenplan.RebuildAll ws, lastRow
    modWochenplan.FastOff

    KnoepfeAktualisieren
    modWochenplan.GotoSheet CTRL_SHEET
    modWochenplan.Info n & " Referenz-Codes sind jetzt feste Werte."
    Exit Sub
Fail:
    modWochenplan.ReportError "Fixieren"
End Sub


Public Sub Fixierung_Aufheben()
    Dim ws As Worksheet, st As Worksheet, fixRw As Long

    Set ws = modWochenplan.WpSheet()
    Set st = modWochenplan.SetSheet()
    If ws Is Nothing Or st Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    If Not modWochenplan.IsPlanFixed() Then
        modWochenplan.Info "Der Plan ist nicht fixiert."
        modWochenplan.GotoSheet CTRL_SHEET
        Exit Sub
    End If

    If modWochenplan.Frage("ALLE Referenz-Codes werden neu durchnummeriert. " & _
              "Bereits weitergegebene Codes stimmen danach nicht mehr. Fortfahren?", _
              vbExclamation + vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

    modWochenplan.FastOn
    On Error GoTo Fail

    fixRw = modWochenplan.FixRow()
    st.Cells(fixRw, "B").Value = FIX_NEIN
    st.Cells(fixRw, "C").Value = "Wird über die Schaltfläche im Blatt '" & _
                                 CTRL_SHEET & "' gesetzt"

    modWochenplan.SetStep "Formeln wiederherstellen"
    modWochenplan.RebuildAll ws, modWochenplan.PlanLastRow(ws)
    modWochenplan.FastOff

    KnoepfeAktualisieren
    modWochenplan.GotoSheet CTRL_SHEET
    modWochenplan.Info "Die Fixierung ist aufgehoben."
    Exit Sub
Fail:
    modWochenplan.ReportError "Aufheben der Fixierung"
End Sub
