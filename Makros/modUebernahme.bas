Attribute VB_Name = "modUebernahme"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Daten aus einer bisherigen Mappe uebernehmen
'  ------------------------------------------------------------------
'  DIE UPDATE-STRATEGIE
'    Makros lassen sich nicht ohne Weiteres in eine fremde Mappe
'    hineinbekommen: der Import im VBA-Editor ist Handarbeit, und jeder
'    Weg per Programm (VBProject, PowerShell, eigenes Programm) haengt
'    an einer Sicherheitseinstellung, die man Kolleginnen und Kollegen
'    nicht zumuten kann und die an Schulen oft gesperrt ist.
'
'    Deshalb laeuft es andersherum: NICHT den neuen Code in die alte
'    Mappe, sondern die alten DATEN in die neue Mappe. Die neue Mappe
'    bringt die aktuellen Makros bereits mit; hier drin passiert nur
'    noch, was Excel ohnehin jederzeit erlaubt - eine zweite Mappe
'    oeffnen und Zellen lesen.
'
'    Damit braucht ein Update keinerlei Einstellung am Rechner:
'    neue Mappe oeffnen, eine Schaltflaeche druecken, fertig.
'
'  WAS UEBERNOMMEN WIRD
'    Einstellungen  Schuljahresstart, Stunden pro Woche, verfuegbare
'                   Wochen, der ganze Schulwochen-Kalender mit seinen
'                   Haekchen, die Ferientabelle und die vier
'                   Kopfangaben
'    Lernbereiche   alle Zeilen bis zur Summenzeile
'    Wochenplan     je Planzeile die Spalten E bis K und M
'
'    Formeln, Ferienzeilen, Referenz-Codes und der Blattschutz werden
'    anschliessend neu aufgebaut - die kommen aus den Makros, nicht
'    aus der alten Datei.
'
'  Die alte Mappe wird nur LESEND geoeffnet und nicht veraendert.
'=====================================================================

'  Beschriftungen, an denen die alte Mappe vermessen wird. Sie muessen
'  zu den aelteren Fassungen passen, nicht nur zur aktuellen.
Private Const Q_WPHEAD  As String = "Referenz Code"
Private Const Q_LBHEAD  As String = "Referenzcode"
Private Const Q_WOCHE   As String = "Schulwoche"
Private Const Q_FERVON  As String = "Ferien von"
Private Const Q_SUMME   As String = "Summe"

Private Const SCAN_ZEILEN As Long = 200
Private Const SCAN_SPALTEN As Long = 30


'=====================================================================
'  Einstiegspunkt - haengt an der Schaltflaeche im Blatt "Steuerung"
'=====================================================================
Public Sub Daten_Uebernehmen()
    Dim pfad As String
    Dim nPlan As Long, nLb As Long, nWo As Long, nFer As Long
    Dim altName As String
    Dim offen As String
    Dim warStill As Boolean

    If Not modWochenplan.LayoutReady() Then Exit Sub

    ' --- Schutz vor Datenverlust -------------------------------------
    ' In eine Mappe, in der schon geplant wurde, wird nichts
    ' hineinkopiert. Sonst waere ein Fehlklick teuer.
    If Not MappeIstLeer() Then
        modWochenplan.Info "In dieser Mappe stehen bereits Daten." & vbCrLf & vbCrLf & _
               "Die Übernahme ist nur in einer frischen Mappe gedacht: " & _
               "neue Mappe öffnen, dort übernehmen, dann speichern.", vbOKOnly
        Exit Sub
    End If

    pfad = DateiWaehlen()
    If Len(pfad) = 0 Then Exit Sub

    If StrComp(pfad, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        modWochenplan.Info "Das ist diese Mappe selbst.", vbOKOnly
        Exit Sub
    End If

    nPlan = UebernahmeAusfuehren(pfad, altName, nLb, nWo, nFer)
    If nPlan < 0 Then Exit Sub

    If nPlan = 0 And nLb = 0 Then
        modWochenplan.Info "In """ & altName & """ waren keine Plandaten zu finden." & vbCrLf & vbCrLf & _
               "Ist es wirklich ein Stoffverteilungsplan?", vbOKOnly
        Exit Sub
    End If

    '  Alles Abgeleitete frisch aufbauen - und gleich im Anschluss den
    '  Wochenplan fertig bauen. Ohne diesen zweiten Schritt steht der
    '  Anwender vor einer Tabelle ohne Ferienzeilen und muss von Hand
    '  nachlegen; das war die erste Rueckfrage nach dem ersten Test.
    '
    '  Beide Routinen melden sich normalerweise selbst - drei Fenster
    '  hintereinander will niemand. Deshalb still, mit Wiederherstellen
    '  des vorherigen Zustands (der Selbsttest laeuft ohnehin still).
    '  Ein Fehler geht dabei NICHT verloren: ReportError schreibt ihn
    '  auch im stillen Modus nach LastError.
    warStill = modWochenplan.IsQuiet()
    modWochenplan.ClearLastError
    modWochenplan.SetQuiet True
    modSteuerung.Setup_Stoffverteilungsplan
    modKalender.UW_Und_Ferien_Generieren
    modWochenplan.SetQuiet warStill

    If Len(modWochenplan.LastError()) > 0 Then
        modWochenplan.Info "Die Daten sind angekommen, aber der Neuaufbau ist " & _
               "gescheitert:" & vbCrLf & vbCrLf & modWochenplan.LastError() & vbCrLf & vbCrLf & _
               "Bitte im Blatt """ & CTRL_SHEET & """ auf ""Wochenplan neu aufbauen"" " & _
               "drücken."
        Exit Sub
    End If

    '  Ganz alte Mappen haben weder Ferientabelle noch die vier
    '  Kopfangaben - beides kam erst spaeter dazu. Das faellt sonst
    '  erst auf, wenn der Wochenplan ohne Ferienzeilen dasteht.
    offen = ""
    If nFer = 0 Then
        offen = offen & vbCrLf & _
                "   - Ferientabelle: in """ & SET_SHEET & """ auf ""Ferien aus " & _
                "Kalender vorschlagen"" drücken."
    End If
    If Len(Trim$(CStr(modWochenplan.SetSheet().Cells( _
           modWochenplan.KopfZeileFach(), modWochenplan.KopfCol() + 1).Value))) = 0 Then
        offen = offen & vbCrLf & _
                "   - Kopfangaben Fach / Klasse / Schule / Lehrkraft in """ & _
                SET_SHEET & """ eintragen, dann ""Kopf übernehmen""."
    End If
    If Len(offen) > 0 Then
        offen = vbCrLf & vbCrLf & "Das gab es in der alten Datei noch nicht:" & offen
    End If

    modWochenplan.GotoSheet CTRL_SHEET
    modWochenplan.Info "Übernommen aus """ & altName & """:" & vbCrLf & vbCrLf & _
           "   " & nPlan & " Planzeilen" & vbCrLf & _
           "   " & nLb & " Lernbereiche" & vbCrLf & _
           "   " & nWo & " Schulwochen, " & nFer & " Ferieneinträge" & offen & vbCrLf & vbCrLf & _
           "Der Wochenplan ist bereits neu aufgebaut. Bitte jetzt " & _
           "speichern - am besten unter dem Namen der bisherigen Datei."
End Sub


'=====================================================================
'  Der eigentliche Vorgang - ohne Dateidialog, ohne Meldung, ohne
'  Setup. Genau deshalb getrennt: so kann modSelbsttest die Uebernahme
'  wirklich durchspielen, statt nur die Lesefunktionen zu pruefen.
'  Der Fehler, den der erste Anlauf beim Nutzer hatte, sass auf der
'  Schreibseite - ein Test, der nur liest, haette ihn nicht gefunden.
'
'  Rueckgabe: Zahl der uebernommenen Planzeilen, -1 nach einem Fehler
'  (die Meldung ist dann schon draussen).
'=====================================================================
Public Function UebernahmeAusfuehren(ByVal pfad As String, ByRef altName As String, _
                                     ByRef nLb As Long, ByRef nWo As Long, _
                                     ByRef nFer As Long) As Long
    Dim alt As Workbook

    On Error GoTo Fail

    '  FastOn steht bewusst VOR dem Oeffnen der alten Mappe:
    '  1. FastOn merkt sich den Zustand von ScreenUpdating und stellt
    '     ihn in FastOff wieder her. Wuerde hier vorher von Hand
    '     ScreenUpdating = False gesetzt, merkte sich FastOn genau
    '     dieses False - und der Bildschirm bliebe nach FastOff
    '     eingefroren.
    '  2. FastOn setzt EnableEvents = False. Damit laeuft das
    '     Workbook_Open der alten Mappe nicht an - deren (aeltere)
    '     Makros sollen hier nichts anfassen.
    modWochenplan.FastOn
    Set alt = Workbooks.Open(Filename:=pfad, ReadOnly:=True, UpdateLinks:=0)
    altName = alt.Name

    modWochenplan.SetStep "Einstellungen übernehmen"
    nWo = EinstellungenHolen(alt, nFer)
    modWochenplan.SetStep "Lernbereiche übernehmen"
    nLb = LernbereicheHolen(alt)
    modWochenplan.SetStep "Wochenplan übernehmen"
    UebernahmeAusfuehren = WochenplanHolen(alt)

    '  PlatzSchaffen klont die eine Musterzeile der Vorlage; die
    '  formatierten Leerzeilen, die die Vorlage darunter mitbringt,
    '  rutschen dabei ans Ende. Hier wegraeumen, damit die Tabelle
    '  sofort sauber aussieht und nicht erst nach dem naechsten
    '  "Wochenplan neu aufbauen".
    modWochenplan.SetStep "Leerzeilen entfernen"
    modWochenplan.LeereEndzeilenEntfernen modWochenplan.WpSheet()

    alt.Close SaveChanges:=False
    Set alt = Nothing
    modWochenplan.FastOff

    '  Nachfassen: bleibt die Quelle offen, stehen ihre Makros ab
    '  sofort in der Liste unter Alt+F8 (Excel zeigt dort die Makros
    '  ALLER geoeffneten Mappen, jeweils mit dem Dateinamen davor).
    '  Das sieht aus, als gaebe es plotzlich alles doppelt.
    ZuMachen altName
    Exit Function

Fail:
    '  ReportError ruft selbst FastOff auf - das stellt Bildschirm,
    '  Ereignisse, Berechnung und Blattschutz wieder her.
    On Error Resume Next
    If Not alt Is Nothing Then alt.Close SaveChanges:=False
    On Error GoTo 0
    modWochenplan.ReportError "Übernehmen der Daten"
    UebernahmeAusfuehren = -1
End Function


'  Eine Mappe dieses Namens schliessen, falls sie noch offen ist.
Private Sub ZuMachen(ByVal nm As String)
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks(nm)
    If wb Is Nothing Then Exit Sub
    If StrComp(wb.FullName, ThisWorkbook.FullName, vbTextCompare) = 0 Then Exit Sub
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub


'  Fuer den Selbsttest: darf hier ueberhaupt uebernommen werden?
Public Function ZielIstLeer() As Boolean
    ZielIstLeer = MappeIstLeer()
End Function


'=====================================================================
'  Ist diese Mappe noch unbeschrieben?
'=====================================================================
Private Function MappeIstLeer() As Boolean
    Dim ws As Worksheet, r As Long, letzte As Long
    Dim c As Variant

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Function

    letzte = modWochenplan.PlanLastRow(ws)
    For r = modWochenplan.WP_FIRST_ROW To letzte
        If Not modWochenplan.IsFerienRow(ws, r) Then
            For Each c In Array("F", "G", "H", "I", "M")
                If Len(Trim$(CStr(ws.Cells(r, CStr(c)).Value))) > 0 Then Exit Function
            Next c
        End If
    Next r
    MappeIstLeer = True
End Function


Private Function DateiWaehlen() As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd Is Nothing Then Exit Function
    With fd
        .Title = "Bisherigen Stoffverteilungsplan auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-Arbeitsmappen", "*.xlsm; *.xlsx"
        If Len(ThisWorkbook.Path) > 0 Then .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = -1 Then DateiWaehlen = .SelectedItems(1)
    End With
    On Error GoTo 0
End Function


'=====================================================================
'  Einstellungen
'=====================================================================
Private Function EinstellungenHolen(ByVal alt As Workbook, ByRef nFer As Long) As Long
    Dim qs As Worksheet, zs As Worksheet
    Dim qKal As Long, zKal As Long, qFer As Long, zFer As Long
    Dim i As Long, n As Long
    Dim sp As Long

    Set qs = BlattOder(alt, SET_SHEET)
    Set zs = modWochenplan.SetSheet()
    If qs Is Nothing Or zs Is Nothing Then Exit Function

    ' --- die drei Eckwerte oben links --------------------------------
    WertKopieren qs, zs, "Schuljahresstart", 1
    WertKopieren qs, zs, "Stunden pro Woche", 1
    WertKopieren qs, zs, "verfügbare Wochen insgesamt", 1

    ' --- Kopfangaben (gibt es in älteren Mappen noch nicht) ----------
    WertKopieren qs, zs, LBL_FACH, 1
    WertKopieren qs, zs, LBL_KLASSE, 1
    WertKopieren qs, zs, LBL_SCHULE, 1
    WertKopieren qs, zs, LBL_LEHRER, 1

    ' --- Schulwochen-Kalender ----------------------------------------
    qKal = ZeileMitText(qs, Q_WOCHE, 1) + 1
    zKal = modWochenplan.WeekFirstRow()
    If qKal > 1 And zKal > 0 Then
        i = 0
        Do While i < SCAN_ZEILEN
            If Not IsNumeric(qs.Cells(qKal + i, 1).Value) Then Exit Do
            If Len(Trim$(CStr(qs.Cells(qKal + i, 1).Value))) = 0 Then Exit Do
            ' Nummer, Datum von/bis und die drei Haekchen
            For sp = 1 To 6
                zs.Cells(zKal + i, sp).Value = qs.Cells(qKal + i, sp).Value
            Next sp
            i = i + 1
            n = n + 1
        Loop
    End If

    ' --- Ferientabelle (in ganz alten Mappen nicht vorhanden) --------
    qFer = ZeileMitText(qs, Q_FERVON, 0)
    zFer = modWochenplan.FerFirstRow()
    If qFer > 0 And zFer > 0 Then
        Dim qSp As Long
        qSp = SpalteMitText(qs, Q_FERVON)
        For i = 0 To 19
            If Not IsDate(qs.Cells(qFer + 1 + i, qSp).Value) Then Exit For
            zs.Cells(zFer + i, FER_COL_VON).Value = qs.Cells(qFer + 1 + i, qSp).Value
            zs.Cells(zFer + i, FER_COL_BIS).Value = qs.Cells(qFer + 1 + i, qSp + 1).Value
            zs.Cells(zFer + i, FER_COL_NAME).Value = qs.Cells(qFer + 1 + i, qSp + 2).Value
            nFer = nFer + 1
        Next i
    End If

    EinstellungenHolen = n
End Function


'  Einen beschrifteten Wert uebernehmen: die Beschriftung in der
'  Quelle suchen und den Wert daneben in die Zelle neben derselben
'  Beschriftung im Ziel schreiben. Fehlt sie irgendwo, passiert nichts.
Private Sub WertKopieren(ByVal qs As Worksheet, ByVal zs As Worksheet, _
                         ByVal bez As String, ByVal versatz As Long)
    Dim qz As Long, qsp As Long, zz As Long, zsp As Long

    qz = ZeileMitText(qs, bez, 0)
    If qz = 0 Then Exit Sub
    qsp = SpalteMitText(qs, bez)
    zz = ZeileMitText(zs, bez, 0)
    If zz = 0 Then Exit Sub
    zsp = SpalteMitText(zs, bez)
    If qsp = 0 Or zsp = 0 Then Exit Sub

    If Len(Trim$(CStr(qs.Cells(qz, qsp + versatz).Value))) = 0 Then Exit Sub
    zs.Cells(zz, zsp + versatz).Value = qs.Cells(qz, qsp + versatz).Value
End Sub


'=====================================================================
'  Lernbereiche
'=====================================================================
Private Function LernbereicheHolen(ByVal alt As Workbook) As Long
    Dim q As Worksheet, z As Worksheet
    Dim qKopf As Long, zKopf As Long, i As Long, n As Long
    Dim gebraucht As Long
    Dim sp As Variant, t As String

    Set q = BlattOder(alt, LB_SHEET)
    Set z = ThisWorkbook.Worksheets(LB_SHEET)
    If q Is Nothing Or z Is Nothing Then Exit Function

    qKopf = ZeileMitTextInSpalte(q, 1, Q_LBHEAD)
    zKopf = modWochenplan.LB_HEAD_ROW()
    If qKopf = 0 Or zKopf = 0 Then Exit Function

    '  Erst zaehlen, dann Platz schaffen. Die Vorlage bringt nur eine
    '  Handvoll Zeilen zwischen Ueberschrift und Summenzeile mit -
    '  ohne diesen Schritt wuerde die Summenzeile ueberschrieben.
    gebraucht = 0
    For i = 0 To 199
        t = Trim$(CStr(q.Cells(qKopf + 1 + i, 4).Value))
        If StrComp(t, Q_SUMME, vbTextCompare) = 0 Then Exit For
        If Len(Trim$(CStr(q.Cells(qKopf + 1 + i, 1).Value))) = 0 And Len(t) = 0 Then Exit For
        gebraucht = gebraucht + 1
    Next i
    If gebraucht = 0 Then Exit Function

    LbPlatzSchaffen z, gebraucht

    '  Uebernommen wird nur, was von Hand eingetragen wird:
    '  A Referenzcode, B von, C bis, D Lernbereich, E Stunden lt.
    '  Lehrplan und J Kompetenzerwartungen. Die Rechenspalten F bis I
    '  und die Summenzeile baut modWochenplan.Lernbereiche_Aufbauen
    '  danach aus dem erkannten Layout - die Formeln der alten Mappe
    '  zeigen auf deren Zeilennummern und waeren hier falsch.
    For i = 0 To gebraucht - 1
        For Each sp In Array(1, 2, 3, 4, 5, 10)
            z.Cells(zKopf + 1 + i, CLng(sp)).Value = q.Cells(qKopf + 1 + i, CLng(sp)).Value
        Next sp
        n = n + 1
    Next i
    LernbereicheHolen = n
End Function


'  Genau so viele Zeilen zwischen Ueberschrift und Summenzeile
'  bereitstellen, wie gebraucht werden - zu wenige werden ergaenzt,
'  zu viele entfernt. Danach steht die Summenzeile direkt unter der
'  letzten Datenzeile.
Private Sub LbPlatzSchaffen(ByVal z As Worksheet, ByVal gebraucht As Long)
    Dim zSum As Long, z1 As Long, vorhanden As Long, diff As Long, i As Long

    zSum = modWochenplan.LbSummeRow(z)
    z1 = modWochenplan.LB_FIRST_ROW()
    If zSum = 0 Or z1 = 0 Then Exit Sub

    vorhanden = zSum - z1
    diff = gebraucht - vorhanden
    If diff = 0 Then Exit Sub

    Application.CutCopyMode = False
    If diff > 0 Then
        '  Die erste Datenzeile als Muster einfuegen: Excel zieht dabei
        '  Format, Gueltigkeit und bedingte Formatierung selbst mit und
        '  weitet die Regelbereiche aus.
        For i = 1 To diff
            z.Rows(z1).Copy
            z.Rows(z1 + 1).Insert Shift:=xlDown
        Next i
        Application.CutCopyMode = False
        z.Range(z.Cells(z1 + 1, "A"), z.Cells(z1 + diff, "J")).ClearContents
    Else
        '  Ueberzaehlige Leerzeilen direkt ueber der Summenzeile weg.
        z.Rows((zSum + diff) & ":" & (zSum - 1)).Delete Shift:=xlUp
    End If
    modWochenplan.ResetKopfzeilen
End Sub


'=====================================================================
'  Wochenplan
'=====================================================================
Private Function WochenplanHolen(ByVal alt As Workbook) As Long
    Dim q As Worksheet, z As Worksheet
    Dim qKopf As Long, qLetzte As Long, r As Long, n As Long
    Dim zeile As Long, gebraucht As Long
    Dim sp As Variant

    Set q = BlattOder(alt, WP_SHEET)
    Set z = modWochenplan.WpSheet()
    If q Is Nothing Or z Is Nothing Then Exit Function

    qKopf = ZeileMitTextInSpalte(q, 2, Q_WPHEAD)
    If qKopf = 0 Then Exit Function
    qLetzte = q.Cells(q.Rows.Count, "G").End(xlUp).Row
    If qLetzte <= qKopf Then Exit Function

    ' Erst zaehlen, wie viele Planzeilen gebraucht werden, und die
    ' Tabelle notfalls verlaengern - die neue Mappe bringt nur eine
    ' begrenzte Zahl Zeilen mit.
    For r = qKopf + 1 To qLetzte
        If Not QuelleIstFerienzeile(q, r) Then gebraucht = gebraucht + 1
    Next r
    PlatzSchaffen z, gebraucht

    '  Geschrieben wird in gebraucht aufeinanderfolgende Zeilen ab
    '  WP_FIRST_ROW.
    '
    '  HIER STAND DER FEHLER der ersten Fassung: eine Sicherung
    '  "If zeile > PlanLastRow(z) Then Exit For". PlanLastRow sucht mit
    '  End(xlUp) von unten nach Inhalt - in der noch leeren Zielmappe
    '  findet es nichts und liefert immer WP_FIRST_ROW. Nach der ersten
    '  geschriebenen Zeile war die Bedingung erfuellt und die Schleife
    '  vorbei: es kam genau eine Zeile an.
    '  Die Sicherung ist auch nicht noetig - PlatzSchaffen hat gerade
    '  eben genug Zeilen angelegt, und Ferienzeilen kann es hier nicht
    '  geben, weil die Uebernahme nur in eine leere Mappe laeuft.
    zeile = modWochenplan.WP_FIRST_ROW
    For r = qKopf + 1 To qLetzte
        If Not QuelleIstFerienzeile(q, r) Then
            For Each sp In Array("E", "F", "G", "H", "I", "J", "K", "M")
                z.Cells(zeile, CStr(sp)).Value = q.Cells(r, CStr(sp)).Value
            Next sp
            zeile = zeile + 1
            n = n + 1
        End If
    Next r
    WochenplanHolen = n
End Function


'  Ferienzeile in der QUELLE: erkennbar am Kennzeichen in Spalte N
'  (aeltere Fassungen schrieben dort "FERIEN") oder daran, dass B bis M
'  verbunden sind.
Private Function QuelleIstFerienzeile(ByVal q As Worksheet, ByVal r As Long) As Boolean
    Dim v As String
    On Error Resume Next
    v = UCase$(Trim$(CStr(q.Cells(r, MARK_COL).Value)))
    If v = MARK_TAG Or v = MARK_TAG_ALT Then
        QuelleIstFerienzeile = True
        Exit Function
    End If
    If q.Cells(r, "B").MergeCells Then
        If q.Cells(r, "B").MergeArea.Columns.Count > 4 Then QuelleIstFerienzeile = True
    End If
End Function


'  So viele Planzeilen bereitstellen, wie gebraucht werden.
'
'  Bewusst NICHT ueber modWochenplan.InsertRowsBelow: die Routine hat
'  ihr eigenes FastOn/FastOff. Aufgerufen aus einem laufenden
'  FastOn-Block heraus wuerde ihr FastOff den Blattschutz mitten im
'  Vorgang wieder einschalten - die naechste Schreiboperation liefe
'  dann in Laufzeitfehler 1004.
Private Sub PlatzSchaffen(ByVal z As Worksheet, ByVal gebraucht As Long)
    Dim vorhanden As Long, letzte As Long, fehlt As Long, i As Long
    Dim anker As Long
    Dim cr() As Long

    '  In der leeren Zielmappe liefert PlanLastRow immer WP_FIRST_ROW
    '  (End(xlUp) findet ohne Inhalt nichts), also gilt genau eine
    '  nutzbare Zeile. Das passt: die Vorlage bringt tatsaechlich nur
    '  eine ausformulierte Planzeile mit, alles darunter ist bloss
    '  formatiert. Die formatierten Leerzeilen rutschen beim Einfuegen
    '  nach unten und raeumt spaeter LeereEndzeilenEntfernen weg.
    letzte = modWochenplan.PlanLastRow(z)
    vorhanden = modWochenplan.ContentRows(z, letzte, cr)
    If vorhanden = 0 Then Exit Sub
    fehlt = gebraucht - vorhanden
    If fehlt <= 0 Then Exit Sub

    ' Die letzte Planzeile so oft als "kopierte Zellen" einfuegen, wie
    ' Zeilen fehlen - dabei uebernimmt Excel selbst Formate,
    ' Kontrollkaestchen, Zeilenhoehe, Gueltigkeit und bedingte
    ' Formatierung und dehnt die Regelbereiche mit aus.
    anker = cr(vorhanden)
    Application.CutCopyMode = False
    For i = 1 To fehlt
        z.Rows(anker).Copy
        z.Rows(anker + 1).Insert Shift:=xlDown
    Next i
    Application.CutCopyMode = False

    z.Range(z.Cells(anker + 1, "A"), z.Cells(anker + fehlt, "V")).ClearContents
End Sub


'=====================================================================
'  Kleine Sucher - arbeiten auf einem BELIEBIGEN Blatt, also auch auf
'  dem einer fremden Mappe.
'=====================================================================
Private Function BlattOder(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    On Error Resume Next
    Set BlattOder = wb.Worksheets(nm)
End Function


'  Erste Zeile, in der irgendwo im Suchbereich dieser Text beginnt.
'  nurSpalte > 0 schraenkt auf eine Spalte ein.
Private Function ZeileMitText(ByVal ws As Worksheet, ByVal txt As String, _
                              ByVal nurSpalte As Long) As Long
    Dim r As Long, c As Long, c1 As Long, c2 As Long
    Dim n As Long, low As String

    n = Len(txt)
    low = LCase$(txt)
    If nurSpalte > 0 Then
        c1 = nurSpalte: c2 = nurSpalte
    Else
        c1 = 1: c2 = SCAN_SPALTEN
    End If

    On Error Resume Next
    For r = 1 To SCAN_ZEILEN
        For c = c1 To c2
            If LCase$(Left$(Trim$(CStr(ws.Cells(r, c).Value)), n)) = low Then
                ZeileMitText = r
                Exit Function
            End If
        Next c
    Next r
End Function


Private Function SpalteMitText(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r As Long, c As Long, n As Long, low As String
    n = Len(txt)
    low = LCase$(txt)
    On Error Resume Next
    For r = 1 To SCAN_ZEILEN
        For c = 1 To SCAN_SPALTEN
            If LCase$(Left$(Trim$(CStr(ws.Cells(r, c).Value)), n)) = low Then
                SpalteMitText = c
                Exit Function
            End If
        Next c
    Next r
End Function


Private Function ZeileMitTextInSpalte(ByVal ws As Worksheet, ByVal spalte As Long, _
                                      ByVal txt As String) As Long
    ZeileMitTextInSpalte = ZeileMitText(ws, txt, spalte)
End Function
