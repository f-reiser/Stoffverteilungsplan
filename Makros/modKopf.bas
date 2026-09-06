Attribute VB_Name = "modKopf"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Titelzeile und PDF-Ausgabe
'  ------------------------------------------------------------------
'  TITELZEILE
'    Eine einzige Zeile aus normalen Zellen oberhalb der
'    Ueberschriftenzeile, dazu eine schmale Abstandszeile. Die Zeile
'    wird in drei verbundene Bloecke geteilt - anders bekommt man in
'    einer Zeile keine drei Ausrichtungen:
'        links   das Schullogo (ohne Logo: der Schulname)
'        Mitte   Fach und Klasse, zentriert
'        rechts  Lehrkraft und Schuljahr
'    Der Schulname steht bewusst nicht als Text da - er steht im Logo.
'
'  LOGOS
'    Die Logos liegen als Bilder FEST IN DER MAPPE, je einmal in
'    "Wochenplan" und "Lernbereiche", benannt wpKopfLogo_<n> mit n =
'    Zeile der Schulliste. Kopf_Aktualisieren blendet nur das zur
'    gewaehlten Schule passende ein und zieht es auf Groesse.
'
'    Bewusst KEIN Kopieren einer Vorlage mehr: Copy/Paste einer Grafik
'    rastert sie in der aktuellen Anzeigegroesse neu - dabei ist am
'    03.09.2026 aus einem 714x80-Logo ein unlesbares 66x55-Bild
'    geworden. Ein bereits eingebettetes Bild nur ein- und auszublenden
'    und zu skalieren kostet dagegen nichts an Qualitaet, weil Excel
'    immer das Original speichert.
'
'  PDF
'    Papierformat kommt aus der Auswahlliste im Blatt "Steuerung"
'    (Name wpPdfFormat), Ziel und Dateiname aus dem normalen
'    Windows-Speichern-Dialog. Keine Rueckfragen per Meldungsfenster.
'=====================================================================

'  Titelzeile + Abstandszeile
Public Const KOPF_ZEILEN As Long = 2
Private Const H_TITEL As Double = 34
Private Const H_LEER  As Double = 6

'  Bilder in "Wochenplan" und "Lernbereiche": wpKopfLogo_1, _2, ...
Public Const LOGO_PRAEFIX As String = "wpKopfLogo_"
Private Const LOGO_HOEHE  As Double = 26
Private Const LOGO_RAND   As Double = 10

'  Breite (in Punkt), die rechts fuer Lehrkraft und Schuljahr
'  reserviert wird. Links richtet sich nach der Breite des Logos.
Private Const ZONE_RECHTS_PT As Double = 200

'  Wie viele Zeilen oberhalb der Ueberschrift hoechstens als alter
'  Titelblock eingesammelt und ersetzt werden.
Private Const KOPF_MAX_ALT As Long = 8

'  Schriftgroesse - fuer ALLE drei Bloecke gleich.
Private Const KOPF_SCHRIFT As Double = 14

'  Spalten, die im PDF auf die Seitenbreite skaliert werden.
Public Const PDF_WP_SPALTE_ERSTE  As String = "B"
Public Const PDF_WP_SPALTE_LETZTE As String = "M"
Public Const PDF_LB_SPALTE_ERSTE  As String = "A"
Public Const PDF_LB_SPALTE_LETZTE As String = "J"

'  Name der Zelle im Blatt "Steuerung", in der das Papierformat steht.
Public Const PDF_FORMAT_NAME As String = "wpPdfFormat"

'  Zuletzt benutzter Ausgabeordner. Steckt NICHT in einer Zelle,
'  sondern als versteckter benannter Wert in der Mappe - damit ist er
'  dauerhaft gespeichert, ohne irgendwo im Blatt aufzutauchen.
Public Const PDF_ORDNER_NAME As String = "wpPdfOrdner"


'=====================================================================
'  Titelzeile setzen
'=====================================================================
Public Sub Kopf_Aktualisieren(Optional ByVal leise As Boolean = False)
    Dim st As Worksheet, ws As Worksheet, lb As Worksheet
    Dim fach As String, klasse As String, schule As String, lehrer As String
    Dim sj As String, logoIdx As Long

    Set st = modWochenplan.SetSheet()
    If st Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    If modWochenplan.KopfZeileFach() = 0 Then
        If leise Then Exit Sub
        modWochenplan.Info "Im Blatt '" & SET_SHEET & "' fehlen die Kopfangaben " & _
               "Fach, Klasse, Schule und Lehrkraft.", vbExclamation
        Exit Sub
    End If

    fach = KopfWert(st, modWochenplan.KopfZeileFach())
    klasse = KopfWert(st, modWochenplan.KopfZeileKlasse())
    schule = KopfWert(st, modWochenplan.KopfZeileSchule())
    lehrer = KopfWert(st, modWochenplan.KopfZeileLehrer())
    sj = Schuljahr(st)
    logoIdx = SchulZeileIndex(st, schule)

    modWochenplan.FastOn
    On Error GoTo Fail

    Set ws = modWochenplan.WpSheet()
    Set lb = Blatt(LB_SHEET)

    modWochenplan.SetStep "Titelzeile " & WP_SHEET
    KopfBauen ws, PDF_WP_SPALTE_ERSTE, PDF_WP_SPALTE_LETZTE, _
              modWochenplan.WP_HEAD_ROW(), fach, klasse, schule, lehrer, sj, logoIdx
    modWochenplan.ResetKopfzeilen

    modWochenplan.SetStep "Titelzeile " & LB_SHEET
    KopfBauen lb, PDF_LB_SPALTE_ERSTE, PDF_LB_SPALTE_LETZTE, _
              modWochenplan.LB_HEAD_ROW(), fach, klasse, schule, lehrer, sj, logoIdx
    modWochenplan.ResetKopfzeilen

    modWochenplan.FastOff
    ' Nur bei einem Klick des Anwenders zurueck aufs Einstellungsblatt.
    ' Im leisen Modus laeuft das aus der Einrichtung heraus, und die
    ' bestimmt selbst, wo der Anwender danach landet.
    If Not leise Then modWochenplan.GotoSheet SET_SHEET
    Exit Sub
Fail:
    modWochenplan.ReportError "Setzen der Titelzeile"
End Sub


'---------------------------------------------------------------------
'  Titelzeile eines Blattes aufbauen.
'
'  Der Block wird bei JEDEM Lauf komplett neu gemacht: erst die Zeilen
'  oberhalb der Ueberschrift einsammeln, die zum Block gehoeren, dann
'  loeschen, dann frisch einfuegen. Das ist von sich aus wiederholbar
'  und raeumt aeltere Fassungen mit auf.
'---------------------------------------------------------------------
Private Sub KopfBauen(ByVal ws As Worksheet, ByVal colFirst As String, _
                      ByVal colLast As String, ByVal headRow As Long, _
                      ByVal fach As String, ByVal klasse As String, _
                      ByVal schule As String, ByVal lehrer As String, _
                      ByVal sj As String, ByVal logoIdx As Long)
    Dim alt As Long, tz As Long, lz As Long
    Dim cErste As Long, cLetzte As Long, cLogo As Long, cRechts As Long
    Dim titel As String, rechts As String, links As String
    Dim logo As Shape, logoBreite As Double

    If ws Is Nothing Then Exit Sub
    If headRow < 1 Then Exit Sub

    cErste = ws.Cells(1, colFirst).Column
    cLetzte = ws.Cells(1, colLast).Column
    If cLetzte <= cErste Then Exit Sub

    '--- alten Block entfernen ---------------------------------------
    ' Die Logos ZUERST freistellen. Sie haengen als Zellanker in der
    ' Titelzeile; wird diese Zeile geloescht, nimmt Excel eine mit der
    ' Zelle verbundene Grafik unter Umstaenden mit. Freigestellte
    ' Formen ueberleben das Loeschen und werden unten neu gesetzt.
    LogosFreistellen ws
    alt = BlockHoehe(ws, headRow, colFirst)

    ' Von Hand verschobene Blockgrenzen merken. Wer die Titelzeile
    ' einmal so aufgeteilt hat, wie er sie haben will, soll das nicht
    ' bei jedem "Kopf uebernehmen" wieder verlieren - aufgefrischt
    ' wird der TEXT, nicht die Aufteilung.
    If alt > 0 Then ZonenLesen ws, headRow - alt, cErste, cLetzte, cLogo, cRechts

    If alt > 0 Then
        ' Fremde Formen ueber den Daten gehoeren zum alten Kopf und
        ' muessen mit weg - sonst liegen zwei Koepfe uebereinander.
        ' Die eingebetteten Logos und die Zeilen-Schaltflaechen sind
        ' davon ausgenommen.
        FormenUeberDenDatenLoeschen ws, headRow - 1
        ws.Rows((headRow - alt) & ":" & (headRow - 1)).Delete Shift:=xlUp
        headRow = headRow - alt
    End If

    '--- Platz schaffen ----------------------------------------------
    ws.Rows(headRow & ":" & (headRow + KOPF_ZEILEN - 1)).Insert Shift:=xlDown
    tz = headRow
    lz = headRow + 1
    headRow = headRow + KOPF_ZEILEN

    ws.Rows(tz).RowHeight = H_TITEL
    ws.Rows(lz).RowHeight = H_LEER
    ZeileNeutral ws, lz, cErste, cLetzte

    '--- Logo waehlen und auf Groesse bringen -------------------------
    ' Zuerst, weil die linke Zone sich nach seiner Breite richtet.
    Set logo = LogoEinstellen(ws, logoIdx)
    If logo Is Nothing Then
        logoBreite = 0
        links = Trim$(schule)
    Else
        logoBreite = logo.Width
        links = ""
    End If

    '--- Texte --------------------------------------------------------
    titel = Trim$(fach)
    If Len(Trim$(klasse)) > 0 Then
        If Len(titel) > 0 Then titel = titel & "   -   "
        titel = titel & "Klasse " & Trim$(klasse)
    End If

    rechts = Trim$(lehrer)
    If Len(rechts) > 0 Then rechts = rechts & "        "
    rechts = rechts & "Schuljahr " & sj

    '--- Zonen bestimmen ----------------------------------------------
    ' Nur wenn keine brauchbaren Grenzen aus dem alten Block kamen:
    ' dann nach BREITE, nicht nach Spaltenzahl - die Spalten sind in
    ' "Wochenplan" und "Lernbereiche" voellig unterschiedlich breit.
    If cLogo <= cErste Or cRechts <= cLogo Or cRechts > cLetzte Then
        If logoBreite > 0 Then
            cLogo = SpalteBeiBreite(ws, cErste, cLetzte, logoBreite + 2 * LOGO_RAND)
        Else
            cLogo = SpalteBeiBreite(ws, cErste, cLetzte, 160)
        End If
        cRechts = SpalteVonRechts(ws, cErste, cLetzte, ZONE_RECHTS_PT)
    End If

    If cRechts <= cLogo + 1 Then
        ' Zu schmal fuer drei Bloecke: links das Logo, rechts alles.
        cLogo = cErste
        cRechts = cLogo + 1
        If cRechts > cLetzte Then cRechts = cLetzte
        Zone ws, tz, cErste, cLogo, links, xlLeft
        Zone ws, tz, cRechts, cLetzte, TrimJoin(titel, rechts), xlRight
    Else
        Zone ws, tz, cErste, cLogo, links, xlLeft
        Zone ws, tz, cLogo + 1, cRechts - 1, titel, xlCenter
        Zone ws, tz, cRechts, cLetzte, rechts, xlRight
    End If

    Rahmen ws, tz, cErste, cLetzte
    If Not logo Is Nothing Then LogoStellen logo, ws, tz, cErste
    FensterFixieren ws, headRow
End Sub


'---------------------------------------------------------------------
'  Farbiger Rahmen um die ganze Titelzeile - derselbe Blauton wie die
'  Ueberschriftenzeile der Tabelle.
'---------------------------------------------------------------------
Private Sub Rahmen(ByVal ws As Worksheet, ByVal r As Long, _
                   ByVal c1 As Long, ByVal c2 As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Borders.LineStyle = xlNone
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, _
                      Color:=modWochenplan.FARBE_KOPF
    End With
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------
'  Die Blockgrenzen aus der vorhandenen Titelzeile ablesen: die linke
'  Zone endet dort, wo der Verbund ab colFirst aufhoert, die rechte
'  beginnt dort, wo der Verbund bis colLast anfaengt.
'---------------------------------------------------------------------
Private Sub ZonenLesen(ByVal ws As Worksheet, ByVal tz As Long, _
                       ByVal cErste As Long, ByVal cLetzte As Long, _
                       ByRef cLogo As Long, ByRef cRechts As Long)
    Dim z As Range

    cLogo = 0: cRechts = 0
    If tz < 1 Then Exit Sub

    On Error Resume Next
    Set z = ws.Cells(tz, cErste)
    If z.MergeCells Then cLogo = z.MergeArea.Column + z.MergeArea.Columns.Count - 1

    Set z = ws.Cells(tz, cLetzte)
    If z.MergeCells Then
        cRechts = z.MergeArea.Column
    Else
        cRechts = cLetzte
    End If
    On Error GoTo 0

    ' Nur uebernehmen, wenn dazwischen noch Platz fuer die Mitte ist.
    If cLogo <= cErste Or cRechts <= cLogo + 1 Or cRechts > cLetzte Then
        cLogo = 0: cRechts = 0
    End If
End Sub


'---------------------------------------------------------------------
'  Wie viele Zeilen ueber der Ueberschrift gehoeren zum Titelblock?
'  Dazu zaehlt eine Zeile, die entweder von colFirst aus einzeilig
'  verbunden ist (so entstehen unsere Bloecke) oder vollstaendig leer
'  ist (die Abstandszeile). In einer frischen Mappe kommt 0 heraus.
'---------------------------------------------------------------------
Private Function BlockHoehe(ByVal ws As Worksheet, ByVal headRow As Long, _
                            ByVal colFirst As String) As Long
    Dim r As Long, n As Long

    For n = 1 To KOPF_MAX_ALT
        r = headRow - n
        If r < 1 Then Exit For
        If Not GehoertZumKopf(ws, r, colFirst) Then Exit For
        BlockHoehe = n
    Next n
End Function


Private Function GehoertZumKopf(ByVal ws As Worksheet, ByVal r As Long, _
                                ByVal colFirst As String) As Boolean
    Dim z As Range, m As Range

    On Error Resume Next

    Set z = ws.Cells(r, colFirst)
    If z Is Nothing Then Exit Function

    If z.MergeCells Then
        Set m = z.MergeArea
        If Not m Is Nothing Then
            GehoertZumKopf = (m.Row = r) And (m.Rows.Count = 1) And _
                             (m.Column = z.Column)
        End If
        Exit Function
    End If

    ' Vollstaendig leere Zeile -> Abstandszeile, gehoert dazu.
    ' Bewusst OHNE MergeCells-Abfrage: bei einer nur teilweise
    ' verbundenen Zeile liefert MergeCells Null, und der Vergleich
    ' damit wuerde Fehler 94 ausloesen.
    GehoertZumKopf = (Application.WorksheetFunction.CountA(ws.Rows(r)) = 0)
    On Error GoTo 0
End Function


'  Freie Formen oberhalb der Daten entfernen - der alte Kopf in jeder
'  Bauform, auch ein von Hand gebautes Textfeld. Die eingebetteten
'  Logos bleiben, die braucht die neue Titelzeile.
Private Sub FormenUeberDenDatenLoeschen(ByVal ws As Worksheet, ByVal bisZeile As Long)
    Dim i As Long, z As Range, nm As String

    On Error Resume Next
    If bisZeile < 1 Then Exit Sub
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, Len(LOGO_PRAEFIX)) <> LOGO_PRAEFIX Then
            Set z = Nothing
            Set z = ws.Shapes(i).TopLeftCell
            If Not z Is Nothing Then
                If z.Row <= bisZeile Then ws.Shapes(i).Delete
            End If
        End If
    Next i
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------
'  Einen Block der Titelzeile schreiben. Schrift und Groesse sind in
'  allen drei Bloecken gleich - nur die Ausrichtung unterscheidet sie.
'---------------------------------------------------------------------
Private Sub Zone(ByVal ws As Worksheet, ByVal r As Long, _
                 ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String, _
                 ByVal ausrichtung As Long)
    If c2 < c1 Then Exit Sub

    On Error Resume Next
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .UnMerge
        .ClearContents
        .Merge
        .HorizontalAlignment = ausrichtung
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Font.Size = KOPF_SCHRIFT
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = modWochenplan.FARBE_KOPF
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        If ausrichtung <> xlCenter Then .IndentLevel = 1
    End With
    ws.Cells(r, c1).Value = txt
    On Error GoTo 0
End Sub


'  Abstandszeile: nichts drin, nichts dran.
Private Sub ZeileNeutral(ByVal ws As Worksheet, ByVal r As Long, _
                         ByVal c1 As Long, ByVal c2 As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .UnMerge
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------
'  Spaltengrenzen nach Breite statt nach Spaltenzahl
'---------------------------------------------------------------------
'  Letzte Spalte, bis zu der von c1 an mindestens "punkte" Breite
'  zusammenkommen (mindestens c1, hoechstens c2).
Private Function SpalteBeiBreite(ByVal ws As Worksheet, ByVal c1 As Long, _
                                 ByVal c2 As Long, ByVal punkte As Double) As Long
    Dim c As Long, summe As Double

    SpalteBeiBreite = c1
    For c = c1 To c2
        summe = summe + ws.Columns(c).Width
        SpalteBeiBreite = c
        If summe >= punkte Then Exit Function
    Next c
End Function


'  Erste Spalte, ab der bis c2 mindestens "punkte" Breite uebrig sind.
Private Function SpalteVonRechts(ByVal ws As Worksheet, ByVal c1 As Long, _
                                 ByVal c2 As Long, ByVal punkte As Double) As Long
    Dim c As Long, summe As Double

    SpalteVonRechts = c2
    For c = c2 To c1 Step -1
        summe = summe + ws.Columns(c).Width
        SpalteVonRechts = c
        If summe >= punkte Then Exit Function
    Next c
End Function


Private Function TrimJoin(ByVal a As String, ByVal b As String) As String
    a = Trim$(a): b = Trim$(b)
    If Len(a) = 0 Then
        TrimJoin = b
    ElseIf Len(b) = 0 Then
        TrimJoin = a
    Else
        TrimJoin = a & "        " & b
    End If
End Function


'=====================================================================
'  Logos
'=====================================================================
'  Alle Logos von der Zelle loesen, damit sie das Loeschen der alten
'  Titelzeile ueberstehen.
Private Sub LogosFreistellen(ByVal ws As Worksheet)
    Dim i As Long

    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    For i = 1 To ws.Shapes.Count
        If Left$(ws.Shapes(i).Name, Len(LOGO_PRAEFIX)) = LOGO_PRAEFIX Then
            ws.Shapes(i).Placement = xlFreeFloating
        End If
    Next i
    On Error GoTo 0
End Sub



'  Das zur Schule passende Logo einblenden, alle anderen ausblenden,
'  und es auf Kopfhoehe bringen. Rueckgabe: die sichtbare Form oder
'  Nothing.
Private Function LogoEinstellen(ByVal ws As Worksheet, ByVal logoIdx As Long) As Shape
    Dim i As Long, s As Shape, treffer As Shape

    On Error Resume Next
    If ws Is Nothing Then Exit Function

    For i = 1 To ws.Shapes.Count
        Set s = ws.Shapes(i)
        If Left$(s.Name, Len(LOGO_PRAEFIX)) = LOGO_PRAEFIX Then
            If s.Name = LOGO_PRAEFIX & logoIdx And logoIdx > 0 Then
                Set treffer = s
                s.Visible = msoTrue
            Else
                s.Visible = msoFalse
            End If
        End If
    Next i

    If treffer Is Nothing Then Exit Function

    ' Seitenverhaeltnis sperren und ueber die HOEHE skalieren - die
    ' Breite ergibt sich daraus. Frueher wurde zusaetzlich die Breite
    ' gesetzt, und das hat das Logo platt gedrueckt.
    With treffer
        .Placement = xlFreeFloating
        .LockAspectRatio = msoTrue
        .Height = LOGO_HOEHE
    End With

    Set LogoEinstellen = treffer
End Function


'  Das Logo mittig in die linke Zone der Titelzeile setzen.
Private Sub LogoStellen(ByVal logo As Shape, ByVal ws As Worksheet, _
                        ByVal tz As Long, ByVal cErste As Long)
    On Error Resume Next
    With logo
        .Left = ws.Cells(tz, cErste).Left + LOGO_RAND
        .Top = ws.Rows(tz).Top + (H_TITEL - .Height) / 2
    End With
    On Error GoTo 0
End Sub


'  Steckt im Blatt ein sichtbares Logo? (fuer den Selbsttest)
Public Function LogoVorhanden(ByVal ws As Worksheet) As Boolean
    Dim i As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    For i = 1 To ws.Shapes.Count
        If Left$(ws.Shapes(i).Name, Len(LOGO_PRAEFIX)) = LOGO_PRAEFIX Then
            If ws.Shapes(i).Visible Then
                LogoVorhanden = True
                Exit Function
            End If
        End If
    Next i
End Function


Private Sub FensterFixieren(ByVal ws As Worksheet, ByVal headRow As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(headRow + 1, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub


'=====================================================================
'  PDF-Ausgabe
'=====================================================================
'  Ohne jede Rueckfrage per Meldungsfenster: das Papierformat steht in
'  der Auswahlliste im Blatt "Steuerung", Ordner und Dateiname kommen
'  aus dem normalen Windows-Speichern-Dialog (der warnt auch selbst,
'  bevor er eine vorhandene Datei ersetzt).
Public Sub PDF_Export()
    Dim ws As Worksheet, lb As Worksheet, st As Worksheet
    Dim papier As Long, ziel As String
    Dim fach As String, klasse As String
    Dim versteckt As Collection

    Set ws = modWochenplan.WpSheet()
    Set lb = Blatt(LB_SHEET)
    Set st = modWochenplan.SetSheet()
    If ws Is Nothing Then Exit Sub
    If Not modWochenplan.LayoutReady() Then Exit Sub

    papier = PapierFormat()

    fach = KopfWert(st, modWochenplan.KopfZeileFach())
    klasse = KopfWert(st, modWochenplan.KopfZeileKlasse())
    ziel = ZielDateiWaehlen(fach, klasse)
    If Len(ziel) = 0 Then Exit Sub          ' abgebrochen

    modWochenplan.FastOn
    On Error GoTo Fail

    modWochenplan.SetStep "Seiteneinrichtung " & WP_SHEET
    SeiteFuerPdfEinrichten ws, papier
    If Not lb Is Nothing Then
        modWochenplan.SetStep "Seiteneinrichtung " & LB_SHEET
        SeiteFuerPdfEinrichten lb, papier
    End If

    ' Nur die beiden gewuenschten Blaetter sollen im PDF landen.
    ' Workbook.ExportAsFixedFormat nimmt ALLE sichtbaren Blaetter -
    ' deshalb die uebrigen kurz ausblenden. Das ist der einzige Weg,
    ' der sich nicht auf undokumentiertes Verhalten der Blattauswahl
    ' verlaesst.
    modWochenplan.SetStep "Blaetter ausblenden"
    AndereBlaetterVerbergen versteckt

    modWochenplan.SetStep "PDF schreiben"
    Application.PrintCommunication = True
    ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ziel, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    BlaetterWiederZeigen versteckt
    modWochenplan.FastOff
    modWochenplan.SeitenumbruecheAus
    OrdnerMerken OrdnerAus(ziel)
    Exit Sub
Fail:
    BlaetterWiederZeigen versteckt
    modWochenplan.ReportError "PDF-Export"
End Sub


'  Papierformat aus der Auswahlliste im Blatt "Steuerung".
Private Function PapierFormat() As Long
    Dim t As String

    PapierFormat = xlPaperA4
    On Error Resume Next
    t = UCase$(Trim$(CStr(ThisWorkbook.Names(PDF_FORMAT_NAME) _
                          .RefersToRange.Value)))
    On Error GoTo 0
    If InStr(1, t, "A3") > 0 Then PapierFormat = xlPaperA3
End Function


'  Windows-Speichern-Dialog. Rueckgabe "" bei Abbruch.
Private Function ZielDateiWaehlen(ByVal fach As String, ByVal klasse As String) As String
    Dim vorschlag As String, ordner As String, antwort As Variant

    vorschlag = "Stoffverteilungsplan"
    If Len(Trim$(fach)) > 0 Then vorschlag = vorschlag & "_" & SauberName(fach)
    If Len(Trim$(klasse)) > 0 Then vorschlag = vorschlag & "_" & SauberName(klasse)
    vorschlag = vorschlag & ".pdf"

    ordner = StartOrdner()
    If Len(ordner) > 0 Then vorschlag = ordner & "\" & vorschlag

    On Error Resume Next
    antwort = Application.GetSaveAsFilename( _
                InitialFileName:=vorschlag, _
                FileFilter:="PDF-Datei (*.pdf), *.pdf", _
                Title:="Stoffverteilungsplan als PDF speichern")
    On Error GoTo 0

    ' Abbrechen liefert den Wahrheitswert False, keinen Text.
    If VarType(antwort) = vbBoolean Then Exit Function
    ZielDateiWaehlen = CStr(antwort)
    If LCase$(Right$(ZielDateiWaehlen, 4)) <> ".pdf" Then
        ZielDateiWaehlen = ZielDateiWaehlen & ".pdf"
    End If

    ' SELBST nachfragen, bevor etwas ueberschrieben wird.
    '
    ' Die Rueckfrage des Windows-Dialogs haengt an
    ' Application.DisplayAlerts, und das schalten die Makros beim
    ' Arbeiten ab. Stand es aus einer vorherigen Aktion noch auf
    ' False, verschwindet die Warnung ersatzlos - beim Nutzer ist
    ' genau das passiert, eine fertige Datei war ohne Nachfrage weg.
    ' Auf eine Warnung, die man nicht in der Hand hat, kann man sich
    ' bei so etwas nicht verlassen.
    If Len(Dir$(ZielDateiWaehlen)) > 0 Then
        If modWochenplan.Frage("Die Datei" & vbCrLf & vbCrLf & "   " & _
                   DateiName(ZielDateiWaehlen) & vbCrLf & vbCrLf & _
                   "gibt es dort schon. Ersetzen?", _
                   vbExclamation + vbOKCancel + vbDefaultButton2) <> vbOK Then
            ZielDateiWaehlen = ""
        End If
    End If
End Function


Private Function DateiName(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i = 0 Then DateiName = p Else DateiName = Mid$(p, i + 1)
End Function


Private Function OrdnerAus(ByVal pfad As String) As String
    Dim i As Long
    i = InStrRev(pfad, "\")
    If i > 1 Then OrdnerAus = Left$(pfad, i - 1)
End Function


'---------------------------------------------------------------------
'  Ordner, in dem der Speichern-Dialog aufgehen soll. In dieser
'  Reihenfolge:
'    1. der zuletzt benutzte Ordner, wenn es ihn noch gibt
'    2. der Ordner der Mappe
'    3. der Standardordner von Excel (nur als letzter Ausweg)
'---------------------------------------------------------------------
Private Function StartOrdner() As String
    Dim p As String

    p = GemerkterOrdner()
    If Len(p) > 0 Then
        If Len(Dir$(p, vbDirectory)) > 0 Then
            StartOrdner = p
            Exit Function
        End If
    End If

    p = MappenOrdner()
    If Len(p) > 0 Then
        StartOrdner = p
        Exit Function
    End If

    On Error Resume Next
    StartOrdner = Application.DefaultFilePath
End Function


'---------------------------------------------------------------------
'  Der Ordner der Mappe - auch dann, wenn Excel sie ueber OneDrive
'  geoeffnet hat und ThisWorkbook.Path deshalb eine Webadresse ist.
'
'  Aus
'     https://<mandant>/personal/<nutzer>/Documents/A/B/C
'  wird
'     <OneDrive-Wurzel>\A\B\C
'  Die Wurzel steht in der Umgebungsvariablen OneDriveCommercial
'  (dienstlich) bzw. OneDrive (privat).
'---------------------------------------------------------------------
Private Function MappenOrdner() As String
    Dim p As String, i As Long, rest As String, wurzel As String, kandidat As String

    p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function
    If LCase$(Left$(p & "    ", 4)) <> "http" Then
        MappenOrdner = p
        Exit Function
    End If

    i = InStr(1, p, "/Documents/", vbTextCompare)
    If i = 0 Then Exit Function
    rest = Mid$(p, i + Len("/Documents/"))
    rest = Replace(rest, "/", "\")
    Do While Right$(rest, 1) = "\"
        rest = Left$(rest, Len(rest) - 1)
    Loop
    If Len(rest) = 0 Then Exit Function

    On Error Resume Next
    wurzel = Environ$("OneDriveCommercial")
    If Len(wurzel) = 0 Then wurzel = Environ$("OneDrive")
    If Len(wurzel) = 0 Then Exit Function

    kandidat = wurzel & "\" & rest
    If Len(Dir$(kandidat, vbDirectory)) > 0 Then MappenOrdner = kandidat
    On Error GoTo 0
End Function


'  Zuletzt benutzten Ordner merken bzw. auslesen. Abgelegt als
'  versteckter benannter Wert - kein Platz im Blatt, ueberlebt das
'  Speichern und taucht im Namens-Manager nicht auf.
Private Sub OrdnerMerken(ByVal ordner As String)
    If Len(ordner) = 0 Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Names(PDF_ORDNER_NAME).Delete
    ThisWorkbook.Names.Add Name:=PDF_ORDNER_NAME, _
                           RefersTo:="=""" & ordner & """", Visible:=False
    On Error GoTo 0
End Sub


Private Function GemerkterOrdner() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(PDF_ORDNER_NAME).RefersTo
    On Error GoTo 0
    ' Gespeichert ist die Form   ="C:\Pfad"
    If Len(s) > 3 Then
        If Left$(s, 2) = "=""" And Right$(s, 1) = """" Then
            GemerkterOrdner = Mid$(s, 3, Len(s) - 3)
        End If
    End If
End Function


Private Sub AndereBlaetterVerbergen(ByRef merker As Collection)
    Dim ws As Worksheet

    Set merker = New Collection
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> WP_SHEET And ws.Name <> LB_SHEET Then
            If ws.Visible = xlSheetVisible Then
                merker.Add ws.Name
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    On Error GoTo 0
End Sub


Private Sub BlaetterWiederZeigen(ByRef merker As Collection)
    Dim i As Long

    If merker Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To merker.Count
        ThisWorkbook.Worksheets(merker(i)).Visible = xlSheetVisible
    Next i
    On Error GoTo 0
    Set merker = Nothing
End Sub


'  Seiteneinrichtung eines Blattes, ohne zu exportieren. Getrennt vom
'  Export, damit der Selbsttest sie pruefen kann, ohne eine Datei zu
'  hinterlassen.
Public Sub SeiteFuerPdfEinrichten(ByVal ws As Worksheet, ByVal papier As Long)
    Dim headRow As Long
    Dim colFirst As String, colLast As String

    If ws Is Nothing Then Exit Sub

    If StrComp(ws.Name, LB_SHEET, vbTextCompare) = 0 Then
        colFirst = PDF_LB_SPALTE_ERSTE
        colLast = PDF_LB_SPALTE_LETZTE
        headRow = modWochenplan.LB_HEAD_ROW()
    Else
        colFirst = PDF_WP_SPALTE_ERSTE
        colLast = PDF_WP_SPALTE_LETZTE
        headRow = modWochenplan.WP_HEAD_ROW()
    End If

    SeiteEinrichten ws, papier, colFirst, colLast, headRow
End Sub


Private Sub SeiteEinrichten(ByVal ws As Worksheet, ByVal papier As Long, _
                            ByVal colFirst As String, ByVal colLast As String, _
                            ByVal headRow As Long)
    Dim lastRow As Long

    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, colFirst).End(xlUp).Row
    If lastRow < headRow Then lastRow = headRow

    ' Kopf- und Fusszeilen ZUERST und mit eingeschalteter
    ' PrintCommunication leeren. Unter PrintCommunication = False sind
    ' diese Zuweisungen am 03.09.2026 nicht angekommen - im PDF stand
    ' danach noch die alte Kopfzeile mit &Z, also dem vollstaendigen
    ' Pfad der Mappe. Das ist nichts, was in einem Plan fuer Kollegen
    ' stehen darf.
    Application.PrintCommunication = True
    With ws.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "&8Seite &P von &N"
        .RightFooter = ""
    End With

    ' Der Rest ist teuer: jede einzelne Eigenschaft wuerde mit dem
    ' Druckertreiber abgestimmt.
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$" & colFirst & "$1:$" & colLast & "$" & lastRow
        .Orientation = xlLandscape
        .PaperSize = papier
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        ' Nur die Spaltenueberschrift wiederholen - die Titelzeile
        ' darueber steht bewusst nur auf der ersten Seite.
        .PrintTitleRows = "$" & headRow & ":$" & headRow
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .HeaderMargin = Application.CentimetersToPoints(0.4)
        .FooterMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True

    ' Gegenprobe: ist die Kopfzeile wirklich leer? Wenn nicht, noch
    ' einmal - lieber langsam als mit dem Dateipfad im PDF.
    If Len(ws.PageSetup.LeftHeader) > 0 Or Len(ws.PageSetup.CenterHeader) > 0 _
       Or Len(ws.PageSetup.RightHeader) > 0 Then
        With ws.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
        End With
    End If
    On Error GoTo 0
End Sub


Private Function SauberName(ByVal t As String) As String
    Dim i As Long, ch As String, s As String
    t = Trim$(t)
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If InStr(1, "\/:*?""<>|", ch) > 0 Then
            s = s & "-"
        ElseIf ch = " " Then
            s = s & "_"
        Else
            s = s & ch
        End If
    Next i
    SauberName = s
End Function


'=====================================================================
'  Werte aus dem Blatt "Einstellungen"
'=====================================================================
Private Function KopfWert(ByVal st As Worksheet, ByVal zeile As Long) As String
    On Error Resume Next
    If st Is Nothing Then Exit Function
    If zeile <= 0 Then Exit Function
    KopfWert = Trim$(CStr(st.Cells(zeile, modWochenplan.KopfCol() + 1).Value))
End Function


Private Function Schuljahr(ByVal st As Worksheet) As String
    Dim y As Long
    y = modKalender.SchoolStartYear(st)
    Schuljahr = y & "/" & Format$((y + 1) Mod 100, "00")
End Function


'  Zeilennummer der Schule in der Schulliste = Nummer ihres Logos.
Private Function SchulZeileIndex(ByVal st As Worksheet, ByVal schule As String) As Long
    Dim r As Long
    If st Is Nothing Then Exit Function
    If modWochenplan.SchulRow() = 0 Then Exit Function
    If Len(Trim$(schule)) = 0 Then Exit Function
    For r = modWochenplan.SchulRow() + 1 To modWochenplan.SchulLastRow()
        If LCase$(Trim$(CStr(st.Cells(r, modWochenplan.SchulCol()).Value))) = _
           LCase$(Trim$(schule)) Then
            SchulZeileIndex = r - modWochenplan.SchulRow()
            Exit Function
        End If
    Next r
End Function


Private Function Blatt(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Blatt = ThisWorkbook.Worksheets(nm)
End Function
