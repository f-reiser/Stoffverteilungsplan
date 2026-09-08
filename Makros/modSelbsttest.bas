Attribute VB_Name = "modSelbsttest"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Selbsttest
'  ------------------------------------------------------------------
'  Spielt die riskanten Abläufe der Reihe nach in echtem Excel durch
'  und schreibt einen Bericht als Textdatei neben die Arbeitsmappe.
'
'  Aufruf:  Alt+F8  ->  Selbsttest
'
'  SICHERUNG: Der Test läuft NUR, wenn "Test" im Dateinamen steht.
'  Er verändert die Mappe unterwegs (fixiert, schiebt Zeilen, baut die
'  Unterrichtswochen neu auf) - deshalb gehört er in eine Kopie und
'  niemals in den echten Plan.
'
'  Kernidee sind die Rundläufe: eine Aktion und ihr Gegenstück müssen
'  den Plan Zeichen für Zeichen so hinterlassen, wie er vorher war.
'  Damit fallen auch Fehler auf, die keine Fehlermeldung erzeugen.
'=====================================================================

Private Const BERICHT_DATEI As String = "Selbsttest_Bericht.txt"

Private mRep As String
Private mOk As Long
Private mFail As Long
Private mFehlerListe As String
Private mErwarte As String
Private mGetroffen As Boolean

'  --- Mutationstest -------------------------------------------------
'  Diese drei gehoeren inhaltlich zu Selbsttest_Pruefen weiter unten
'  und standen deshalb zuerst dort - direkt vor der Prozedur.
'  DAS GEHT IN VBA NICHT: modulweite Deklarationen duerfen
'  ausschliesslich im Deklarationsteil ganz oben stehen. Zwischen zwei
'  Prozeduren quittiert Excel das mit "Nach End Sub, End Function oder
'  End Property koennen nur Kommentare stehen" - und in der Folge mit
'  "Variable nicht definiert" an jeder Verwendung.
Private Const MUT_ANZAHL As Long = 10
Private mSaveText As String
Private mSaveZahl As Double
Private mSaveName As String


'=====================================================================
'  Einstiegspunkt
'=====================================================================
Public Sub Selbsttest()
    Dim ws As Worksheet, pfad As String, t0 As Single
    Dim errNr As Long, errTxt As String, errSchritt As String

    If InStr(1, ThisWorkbook.Name, "test", vbTextCompare) = 0 Then
        MsgBox "Der Selbsttest läuft nur in einer Kopie." & vbCrLf & vbCrLf & _
               "Er verändert die Mappe unterwegs: er fixiert, schiebt Zeilen hin " & _
               "und her und baut die Unterrichtswochen neu auf." & vbCrLf & vbCrLf & _
               "Bitte die Datei unter einem Namen speichern, der ""Test"" enthält " & _
               "(z. B. Stoffverteilungsplan_Test.xlsm), und den Test dort starten.", _
               vbOKOnly, "Selbsttest"
        Exit Sub
    End If

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then
        MsgBox "Blatt '" & WP_SHEET & "' nicht gefunden.", vbOKOnly, "Selbsttest"
        Exit Sub
    End If

    mRep = "": mOk = 0: mFail = 0: mFehlerListe = ""
    t0 = Timer

    Zeile "=========================================================="
    Zeile " Selbsttest Stoffverteilungsplan"
    Zeile " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Zeile " Datei  : " & ThisWorkbook.Name
    Zeile " Excel  : " & Application.Version & " (Build " & Application.Build & ")"
    Zeile "=========================================================="

    modWochenplan.SetQuiet True
    On Error GoTo Fail

    T0_Erreichbarkeit ws
    T1_Aufbau ws
    T1b_Zustand ws
    T2_Aktualisieren ws
    T3_Wochen ws
    T3b_Aufraeumen ws
    T4_Verschieben ws
    T5_EinfuegenLoeschen ws
    T6_Fixieren ws
    T7_Blattschutz ws
    T8_Kopf ws
    T9_Fokus ws
    T10_Uebernahme ws

    modWochenplan.SetQuiet False

    Zeile ""
    Zeile "=========================================================="
    Zeile " Ergebnis:  " & mOk & " bestanden,  " & mFail & " durchgefallen"
    Zeile " Dauer   :  " & Format$(Timer - t0, "0.0") & " s"
    Zeile "=========================================================="

    pfad = BerichtSchreiben()

    MsgBox IIf(mFail = 0, "Selbsttest bestanden.", "Selbsttest MIT FEHLERN.") & vbCrLf & vbCrLf & _
           mOk & " Prüfungen bestanden, " & mFail & " durchgefallen." & vbCrLf & _
           "Dauer: " & Format$(Timer - t0, "0.0") & " Sekunden." & _
           IIf(mFail = 0, "", vbCrLf & vbCrLf & "Durchgefallen:" & mFehlerListe) & _
           vbCrLf & vbCrLf & "Bericht:" & vbCrLf & pfad, _
           vbOKOnly, "Selbsttest"
    Exit Sub

Fail:
    ' Nummer und Text SOFORT sichern: BerichtSchreiben setzt das
    ' Err-Objekt zurueck, danach staende in der Meldung nur "Fehler 0".
    errNr = Err.Number
    errTxt = Err.Description
    errSchritt = modWochenplan.CurrentStep()

    modWochenplan.SetQuiet False
    modWochenplan.FastOff
    Zeile ""
    Zeile "*** ABBRUCH: Laufzeitfehler " & errNr & " - " & errTxt
    Zeile "*** Schritt: " & errSchritt
    Zeile ""
    Zeile " Bis hierher: " & mOk & " bestanden, " & mFail & " durchgefallen"
    pfad = BerichtSchreiben()
    MsgBox "Der Selbsttest ist abgebrochen." & vbCrLf & vbCrLf & _
           "Fehler " & errNr & ": " & errTxt & vbCrLf & _
           "Schritt: " & errSchritt & vbCrLf & vbCrLf & _
           "Bis hierher: " & mOk & " bestanden, " & mFail & " durchgefallen." & vbCrLf & vbCrLf & _
           "Bericht:" & vbCrLf & pfad, vbOKOnly, "Selbsttest"
End Sub


'=====================================================================
'  0  Erreichbarkeit trotz Option Private Module
'  ------------------------------------------------------------------
'  Seit 05.09.2026 tragen alle Module ausser modStart die Zeile
'  "Option Private Module". Das kuerzt die Makroliste unter Alt+F8 von
'  30 auf vier Eintraege - haengt aber daran, dass Shape.OnAction die
'  Makros trotzdem findet. Microsoft dokumentiert nur, dass die
'  Public-Teile "innerhalb des Projekts weiterhin verfuegbar" bleiben;
'  ueber OnAction steht dort nichts. Deshalb wird es hier gemessen
'  statt vermutet - und zwar als ERSTER Abschnitt, damit ein Fehlschlag
'  nicht unter hundert anderen Zeilen verschwindet.
'=====================================================================
Private Sub T0_Erreichbarkeit(ByVal ws As Worksheet)
    Dim ok As Boolean, i As Long, s As Shape
    Dim wsCtrl As Worksheet, ohne As Long, gefunden As Long
    Dim alle As String, doppelt As Long, doppelName As String

    Abschnitt "0  Erreichbarkeit trotz Option Private Module"

    ' --- 1. Application.Run in ein Modul mit Option Private Module ---
    On Error Resume Next
    Err.Clear
    Application.Run "modWochenplan.SelbsttestProbe"
    ok = (Err.Number = 0)
    If Not ok Then Notiz "Fehler " & Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    Chk "Makro in einem Modul mit Option Private Module ist ueber seinen " & _
        "Namen erreichbar", ok And modWochenplan.ProbeGelaufen()

    ' --- 2. Der Weg ueber modStart (dort steht die Zeile NICHT) ------
    On Error Resume Next
    Err.Clear
    Application.Run "modStart.Probe_modStart"
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    Chk "Gegenprobe: modStart (ohne Option Private Module) ist erreichbar", ok

    ' --- 3. Jede Schaltflaeche hat ein Ziel -------------------------
    Set wsCtrl = Blatt(CTRL_SHEET)
    If Not wsCtrl Is Nothing Then
        alle = "|"
        For i = 1 To wsCtrl.Shapes.Count
            Set s = wsCtrl.Shapes(i)
            If Len(s.OnAction) = 0 Then
                ohne = ohne + 1
            ElseIf InStr(s.OnAction, ".") = 0 Then
                ohne = ohne + 1
            Else
                gefunden = gefunden + 1
                '  Doppelte Ziele mitzaehlen - siehe unten.
                If InStr(1, alle, "|" & s.OnAction & "|", vbTextCompare) > 0 Then
                    doppelt = doppelt + 1
                    If Len(doppelName) = 0 Then doppelName = s.OnAction
                Else
                    alle = alle & s.OnAction & "|"
                End If
            End If
        Next i
        Chk gefunden & " Schaltflaechen in " & CTRL_SHEET & " zeigen auf ein Makro", _
            (ohne = 0), ohne & " ohne modulqualifiziertes Ziel"

        '  --- Der stille Zwilling ---------------------------------
        '  Eine Schaltflaeche, die das Makro einer ANDEREN aufruft,
        '  faellt oben nicht auf: sie hat ja ein Ziel. Genau so ein
        '  Zustand ist am 05.09.2026 entstanden, als die Ruecknahme
        '  einer Mutation das gemerkte Makro der falschen Form
        '  verpasst hat. Zwei Knoepfe kaputt, einer davon unsichtbar.
        '  In diesem Blatt hat jede Schaltflaeche genau ein eigenes
        '  Makro - ein doppeltes Ziel ist immer ein Fehler.
        Chk "Jede Schaltflaeche in " & CTRL_SHEET & " hat ihr eigenes Makro", _
            (doppelt = 0), doppelt & " x doppelt vergeben, zuerst: " & doppelName
    End If

    ' --- 4. Die Zeilen-Schaltflaechen im Wochenplan -----------------
    ohne = 0: gefunden = 0
    For i = 1 To ws.Shapes.Count
        Set s = ws.Shapes(i)
        If Left$(s.Name, 5) = "wpBtn" Then
            If Len(s.OnAction) = 0 Then ohne = ohne + 1 Else gefunden = gefunden + 1
        End If
    Next i
    Chk gefunden & " Zeilen-Schaltflaechen haben ein Makro hinterlegt", (ohne = 0)
End Sub


'=====================================================================
'  1  Aufbau und Layout
'=====================================================================
Private Sub T1_Aufbau(ByVal ws As Worksheet)
    Dim cr() As Long, n As Long, nf As Long, r As Long, lastRow As Long

    Abschnitt "1  Aufbau und Layout"

    Chk "Blatt " & WP_SHEET, Not modWochenplan.WpSheet() Is Nothing
    Chk "Blatt " & SET_SHEET, Not modWochenplan.SetSheet() Is Nothing
    Chk "Blatt " & LB_SHEET, Not Blatt(LB_SHEET) Is Nothing
    Chk "Blatt " & CTRL_SHEET, Not Blatt(CTRL_SHEET) Is Nothing

    Chk "Layout der Einstellungen erkannt", modWochenplan.RefreshLayout(True)
    Notiz modWochenplan.LayoutInfo()

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then nf = nf + 1
    Next r
    Notiz "Letzte Planzeile " & lastRow & ", davon " & n & " Inhaltszeilen und " & _
          nf & " Ferienzeilen"
    Chk "Mindestens 5 Inhaltszeilen vorhanden", n >= 5
    Chk "Fixierung ist " & IIf(modWochenplan.IsPlanFixed(), "aktiv", "offen"), True
End Sub


'=====================================================================
'  1b  Zustand der Mappe - Pruefungen OHNE Nebenwirkung
'  ------------------------------------------------------------------
'  Dieser Abschnitt fasst nichts an. Er beschreibt nur, wie eine
'  gesunde Mappe aussieht. Genau deshalb ist er der Abschnitt, an dem
'  der Mutationstest (Selbsttest_Pruefen) ansetzt: eine Sabotage bleibt
'  bis zur Pruefung stehen, weil hier nichts repariert wird.
'
'  Jede Zeile hier hat einen Anlass in der Vergangenheit:
'    - Rechenspalten und Summenzeile: standen als feste Formeln im
'      Blatt, die Vorlage brachte sogar =SUM(E10:E10) mit.
'    - Update-Blatt: liess sich wegen des Strukturschutzes lautlos
'      nicht loeschen.
'    - Stand der Anleitung: das Blatt blieb stehen und damit sein
'      alter Text.
'    - Leerzeilen unter der Tabelle: 35 Stueck, und der alte Check in
'      Abschnitt 3b hat sie NICHT gesehen (siehe dort).
'=====================================================================
Private Sub T1b_Zustand(ByVal ws As Worksheet)
    Dim zl As Worksheet, st As Worksheet
    Dim z1 As Long, zSum As Long, zLetzte As Long, r As Long
    Dim c As Variant, fehlt As String
    Dim leere As Long, ohneKennzeichen As Long, letzte As Long
    Dim vorher As Long, e As Variant

    Abschnitt "1b  Zustand der Mappe"

    Set zl = modWochenplan.LbSheet()
    Set st = modWochenplan.SetSheet()

    ' --- Lernbereiche: die vier Rechenspalten sind Formeln ------------
    If Not zl Is Nothing Then
        z1 = modWochenplan.LB_FIRST_ROW()
        zSum = modWochenplan.LbSummeRow(zl)
        zLetzte = modWochenplan.LbLastDataRow(zl, zSum)
        fehlt = ""
        For r = z1 To zLetzte
            For Each c In Array("F", "G", "H", "I")
                If Left$(zl.Cells(r, CStr(c)).Formula & " ", 1) <> "=" Then
                    fehlt = fehlt & " " & CStr(c) & r
                End If
            Next c
        Next r
        Chk "Rechenspalten F bis I der " & LB_SHEET & " sind Formeln", _
            (fehlt = ""), "ohne Formel:" & fehlt

        Chk "Summenzeile gefunden", (zSum > 0)
        If zSum > 0 Then
            Chk "Summenzeile steht direkt unter dem letzten Lernbereich", _
                (zSum = zLetzte + 1), "Summe " & zSum & ", letzter Lernbereich " & zLetzte
            Chk "Summenformel deckt genau die Datenzeilen ab", _
                (InStr(1, zl.Cells(zSum, "E").Formula, _
                       "E" & z1 & ":E" & zLetzte) > 0), _
                zl.Cells(zSum, "E").Formula
        End If
    End If

    ' --- Das entfallene Update-Blatt darf nicht mehr da sein ----------
    fehlt = ""
    For Each e In Array("Update", "VBA-Update")
        If Not Blatt(CStr(e)) Is Nothing Then fehlt = fehlt & " " & CStr(e)
    Next e
    Chk "Kein Update-Blatt mehr in der Mappe", (fehlt = ""), "gefunden:" & fehlt

    ' --- Anleitung auf dem aktuellen Stand ----------------------------
    Chk "Blatt " & modAnleitung.SHEET_HELP & " vorhanden", _
        Not Blatt(modAnleitung.SHEET_HELP) Is Nothing
    Chk "Anleitung ist auf dem aktuellen Stand", _
        (modAnleitung.AnleitungStand() = modAnleitung.ANLEITUNG_STAND), _
        "in der Mappe: """ & modAnleitung.AnleitungStand() & _
        """, erwartet: """ & modAnleitung.ANLEITUNG_STAND & """"

    ' --- Blattschutz --------------------------------------------------
    Chk "Blattschutz ist aktiv", modSchutz.SchutzAktiv()

    ' --- Jede verbundene Planzeile traegt ihr Kennzeichen -------------
    '  Ohne das Kennzeichen in Spalte N gilt eine Ferienzeile als
    '  normale Planzeile - der Formelneuaufbau schreibt dann in die
    '  verbundenen Zellen und endet in Laufzeitfehler 1004.
    letzte = modWochenplan.PlanLastRow(ws)
    ohneKennzeichen = 0
    For r = WP_FIRST_ROW To letzte
        If ws.Cells(r, "B").MergeCells Then
            If ws.Cells(r, "B").MergeArea.Columns.Count > 4 Then
                If Not modWochenplan.IsFerienRow(ws, r) Then
                    ohneKennzeichen = ohneKennzeichen + 1
                End If
            End If
        End If
    Next r
    Chk "Jede verbundene Zeile traegt das Kennzeichen in Spalte " & MARK_COL, _
        (ohneKennzeichen = 0), ohneKennzeichen & " ohne Kennzeichen"

    ' --- Leerzeilen UNTERHALB der Tabelle -----------------------------
    '  Der Check in Abschnitt 3b beginnt bei PlanLastRow und sucht nach
    '  OBEN. PlanLastRow sucht mit End(xlUp) nach INHALT und sieht die
    '  nur formatierten Reservezeilen darunter gar nicht - er hat
    '  "keine leeren Zeilen" gemeldet, waehrend 35 Stueck dastanden.
    '  Erkannt werden sie an der Zeilenhoehe: eine Planzeile ist hoch
    '  (45,75), eine Zeile ausserhalb der Tabelle hat Standardhoehe.
    leere = 0
    r = letzte + 1
    Do While r <= letzte + 200
        If ws.Rows(r).RowHeight < 30 Then Exit Do
        If Not NurLeer(ws, r) Then Exit Do
        leere = leere + 1
        r = r + 1
    Loop
    Chk "Keine formatierten Leerzeilen unter der Tabelle", (leere = 0), _
        leere & " Zeile(n) ab Zeile " & (letzte + 1) & " sehen aus wie Planzeilen"

    ' --- Unterrichtswochen streng aufsteigend -------------------------
    vorher = 0
    fehlt = ""
    For r = WP_FIRST_ROW To letzte
        If Not modWochenplan.IsFerienRow(ws, r) Then
            If IsNumeric(ws.Cells(r, "E").Value) And Len(Trim$(CStr(ws.Cells(r, "E").Value))) > 0 Then
                If CLng(ws.Cells(r, "E").Value) < vorher Then
                    If fehlt = "" Then fehlt = "Zeile " & r
                End If
                vorher = CLng(ws.Cells(r, "E").Value)
            End If
        End If
    Next r
    Chk "Unterrichtswochen laufen nirgends rueckwaerts", (fehlt = ""), fehlt
End Sub


'=====================================================================
'  2  Formelneuaufbau darf nichts verändern
'=====================================================================
Private Sub T2_Aktualisieren(ByVal ws As Worksheet)
    Dim vor As String, nach As String, r As Long, lastRow As Long
    Dim cr() As Long, n As Long

    Abschnitt "2  Formeln neu aufbauen"

    vor = Snap(ws)
    modWochenplan.ClearLastError
    modWochenplan.Wochenplan_Aktualisieren
    Chk "Aktualisieren ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    nach = Snap(ws)
    Chk "Aktualisieren lässt alle Werte unverändert", vor = nach

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    If n > 0 Then
        Chk "Inhaltszeile hat Formel in Spalte A", ws.Cells(cr(1), "A").HasFormula
        Chk "Inhaltszeile hat Formel in Spalte C", ws.Cells(cr(1), "C").HasFormula
        Chk "Inhaltszeile hat Formel in Spalte V", ws.Cells(cr(1), "V").HasFormula
        Chk "Auswahlliste in Spalte K vorhanden", HatGueltigkeit(ws.Cells(cr(1), "K"))
    End If

    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then
            Chk "Ferienzeile " & r & " ohne Formel in Spalte C", _
                Not ws.Cells(r, "C").HasFormula
            Chk "Ferienzeile " & r & " ist B:M verbunden", _
                ws.Cells(r, "B").MergeArea.Address(False, False) = "B" & r & ":M" & r
            Exit For
        End If
    Next r
End Sub


'=====================================================================
'  3  Wochenplan neu aufbauen (UW_Und_Ferien_Generieren)
'=====================================================================
Private Sub T3_Wochen(ByVal ws As Worksheet)
    Dim cr() As Long, n As Long, i As Long, lastRow As Long
    Dim aufsteigend As Boolean, letzteUW As Long, ohneUW As Long
    Dim r As Long, nf As Long

    Abschnitt "3  Wochenplan neu aufbauen"

    modWochenplan.ClearLastError
    modKalender.UW_Und_Ferien_Generieren
    Chk "Erzeugen ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    Notiz modWochenplan.LastInfo()

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then nf = nf + 1
    Next r
    Notiz "Jetzt " & n & " Inhaltszeilen und " & nf & " Ferienzeilen"
    Chk "Es sind Ferienzeilen entstanden", nf > 0

    ' Achtung: IsNumeric(Empty) liefert True. Leere Zellen entstehen
    ' voellig regulaer, wenn es mehr Planzeilen als verfuegbare Wochen
    ' gibt - die duerfen die Pruefung nicht durchfallen lassen.
    aufsteigend = True
    letzteUW = 0
    ohneUW = 0
    For i = 1 To n
        If Len(T(ws.Cells(cr(i), "E").Value)) = 0 Then
            ohneUW = ohneUW + 1
        ElseIf IsNumeric(ws.Cells(cr(i), "E").Value) Then
            If CLng(ws.Cells(cr(i), "E").Value) <= letzteUW Then aufsteigend = False
            letzteUW = CLng(ws.Cells(cr(i), "E").Value)
        End If
    Next i
    Chk "Unterrichtswochen laufen streng aufsteigend", aufsteigend
    If ohneUW > 0 Then Notiz ohneUW & " Planzeile(n) ohne Unterrichtswoche " & _
                            "(mehr Zeilen als verfuegbare Wochen) - kein Fehler."

    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then
            Chk "Ferienzeile " & r & " hat Text in Spalte B", _
                Len(T(ws.Cells(r, "B").Value)) > 0
            Chk "Ferienzeile " & r & " hat keine Unterrichtswoche", _
                Len(T(ws.Cells(r, "E").Value)) = 0
        End If
    Next r

    ' O1 meldet inhaltliche Warnungen des PLANS (falscher Lehrplan-Code,
    ' UW nicht verfuegbar ...). Das ist kein Makrofehler, sondern etwas,
    ' das der Nutzer selbst entscheidet - deshalb nur als Notiz.
    If Len(T(ws.Range("O1").Value)) = 0 Then
        Notiz "Warnspalte O1 ist leer - der Plan ist inhaltlich stimmig."
    Else
        Notiz "Hinweis: O1 meldet """ & T(ws.Range("O1").Value) & """ - " & _
              "inhaltliche Warnungen im Plan, kein Makrofehler."
    End If
End Sub


'=====================================================================
'  4  Verschieben über eine Ferienzeile hinweg (Rundlauf)
'=====================================================================
'=====================================================================
'  3b  Tabellenende und Zeilen-Schaltflaechen
'=====================================================================
Private Sub T3b_Aufraeumen(ByVal ws As Worksheet)
    Dim letzte As Long, r As Long, leere As Long
    Dim erste As Long, letzteInhalt As Long

    Abschnitt "3b  Tabellenende und Zeilen-Schaltflaechen"

    letzte = modWochenplan.PlanLastRow(ws)
    erste = WP_FIRST_ROW

    ' --- leere Zeilen am Ende ---------------------------------------
    leere = 0
    For r = letzte To erste + 1 Step -1
        If modWochenplan.IsFerienRow(ws, r) Then Exit For
        If Not NurLeer(ws, r) Then Exit For
        leere = leere + 1
    Next r
    '  ACHTUNG, blinder Fleck: diese Schleife beginnt bei PlanLastRow.
    '  Zeilen UNTERHALB davon sieht sie nicht - genau die 35 nur
    '  formatierten Reservezeilen, die nach dem ersten Import
    '  stehengeblieben sind, hat sie als "ok" durchgewinkt. Der Check
    '  dafuer sitzt in Abschnitt 1b; hier geht es nur um leere Zeilen
    '  INNERHALB des Plans.
    Chk "Keine leeren Zeilen innerhalb des Plans am Ende", leere = 0, _
        leere & " leere Zeile(n) unter der letzten Eintragung"

    ' --- Schaltflaechen am oberen und unteren Rand -------------------
    letzteInhalt = 0
    For r = letzte To erste Step -1
        If Not modWochenplan.IsFerienRow(ws, r) Then
            letzteInhalt = r
            Exit For
        End If
    Next r

    ws.Activate
    modWochenplan.UpdateRowButtons ws.Range("F" & erste)
    Chk "In der ersten Planzeile fehlt ""nach oben""", _
        Not FormSichtbar(ws, "wpBtnUp")
    Chk "In der ersten Planzeile gibt es ""nach unten""", _
        FormSichtbar(ws, "wpBtnDown")

    If letzteInhalt > erste Then
        modWochenplan.UpdateRowButtons ws.Range("F" & letzteInhalt)
        Chk "In der letzten Planzeile fehlt ""nach unten""", _
            Not FormSichtbar(ws, "wpBtnDown")
        Chk "In der letzten Planzeile gibt es ""nach oben""", _
            FormSichtbar(ws, "wpBtnUp")

        ' Der Zeilenbereich muss auch am AUSGEBLENDETEN Knopf stimmen:
        ' "+" und "-" lesen ihn ueber ReadBlock von wpBtnUp ab, und
        ' ein stehengebliebener alter Wert wuerde an der falschen
        ' Stelle einfuegen oder loeschen.
        Chk "Der Zeilenbereich stimmt auch am ausgeblendeten Knopf", _
            FormBereich(ws, "wpBtnDown") = letzteInhalt & ":" & letzteInhalt, _
            "steht auf '" & FormBereich(ws, "wpBtnDown") & "'"
    End If
End Sub


Private Function NurLeer(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array("E", "F", "G", "H", "I", "K", "M")
        If Len(Trim$(CStr(ws.Cells(r, CStr(c)).Value))) > 0 Then Exit Function
    Next c
    NurLeer = True
End Function


Private Function FormSichtbar(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = Nothing
    Set s = ws.Shapes(nm)
    If Not s Is Nothing Then FormSichtbar = s.Visible
End Function


Private Function FormBereich(ByVal ws As Worksheet, ByVal nm As String) As String
    Dim s As Shape
    On Error Resume Next
    Set s = Nothing
    Set s = ws.Shapes(nm)
    If Not s Is Nothing Then FormBereich = s.AlternativeText
End Function


Private Sub T4_Verschieben(ByVal ws As Worksheet)
    Dim cr() As Long, n As Long, lastRow As Long
    Dim i As Long, idx As Long, ferZeile As Long
    Dim vor As String, inhaltVor As String

    Abschnitt "4  Zeilen verschieben (Rundlauf über eine Ferienzeile)"

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)

    ' Inhaltszeile suchen, deren Vorgängerzeile eine Ferienzeile ist
    idx = 0
    For i = 2 To n
        If cr(i) - cr(i - 1) > 1 Then
            idx = i
            ferZeile = cr(i) - 1
            Exit For
        End If
    Next i

    If idx = 0 Then
        Notiz "Übersprungen - keine Ferienzeile zwischen zwei Inhaltszeilen gefunden."
        Exit Sub
    End If
    Notiz "Getestet wird Zeile " & cr(idx) & " über die Ferienzeile " & ferZeile & " hinweg."

    vor = Snap(ws)
    inhaltVor = T(ws.Cells(cr(idx), "G").Value)

    modWochenplan.ClearLastError
    modWochenplan.MoveBlock cr(idx), cr(idx), -1
    Chk "Verschieben nach oben ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    ' MoveBlock steigt in Randfaellen still mit Beep aus. Ohne diese
    ' Pruefung wuerde der Rundlauf unten auch dann bestehen, wenn gar
    ' nichts passiert ist.
    Chk "Es wurde tatsächlich etwas verschoben", Snap(ws) <> vor
    Chk "Die Ferienzeile blieb an ihrem Platz", modWochenplan.IsFerienRow(ws, ferZeile)
    Chk "Der Inhalt ist über die Ferienzeile gesprungen", _
        T(ws.Cells(cr(idx - 1), "G").Value) = inhaltVor, _
        "erwartet in Zeile " & cr(idx - 1) & ": " & inhaltVor

    modWochenplan.ClearLastError
    modWochenplan.MoveBlock cr(idx - 1), cr(idx - 1), 1
    Chk "Verschieben nach unten ohne Fehlermeldung", modWochenplan.LastError() = ""
    Chk "Rundlauf: der Plan ist wieder wie vorher", Snap(ws) = vor
End Sub


'=====================================================================
'  5  Zeile einfügen und wieder löschen (Rundlauf)
'=====================================================================
Private Sub T5_EinfuegenLoeschen(ByVal ws As Worksheet)
    Dim cr() As Long, n As Long, lastRow0 As Long, lastRow1 As Long
    Dim vor As String, r As Long

    Abschnitt "5  Zeile einfügen und löschen (Rundlauf)"

    lastRow0 = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow0, cr)
    If n < 2 Then Notiz "Übersprungen - zu wenige Zeilen.": Exit Sub

    vor = Snap(ws)
    r = cr(2)

    modWochenplan.ClearLastError
    modWochenplan.InsertRowsBelow r, r
    Chk "Einfügen ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    lastRow1 = modWochenplan.PlanLastRow(ws)
    Chk "Es ist genau eine Zeile dazugekommen", lastRow1 = lastRow0 + 1, _
        "vorher " & lastRow0 & ", jetzt " & lastRow1
    Chk "Die neue Zeile ist leer (Spalte G)", Len(T(ws.Cells(r + 1, "G").Value)) = 0
    Chk "Die neue Zeile ist keine Ferienzeile", Not modWochenplan.IsFerienRow(ws, r + 1)
    Chk "Die neue Zeile hat eine Formel in Spalte C", ws.Cells(r + 1, "C").HasFormula

    modWochenplan.ClearLastError
    modWochenplan.DeleteRowsBlock r + 1, r + 1
    Chk "Löschen ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    Chk "Die Zeilenzahl stimmt wieder", modWochenplan.PlanLastRow(ws) = lastRow0
    Chk "Rundlauf: der Plan ist wieder wie vorher", Snap(ws) = vor
End Sub


'=====================================================================
'  6  Fixieren, Code-Vergabe, Fixierung aufheben
'=====================================================================
Private Sub T6_Fixieren(ByVal ws As Worksheet)
    Dim cr() As Long, n As Long, lastRow As Long
    Dim warFixiert As Boolean, i As Long, idx As Long
    Dim codeVor As String, vor As String, neuerCode As String, naechste As Long

    Abschnitt "6  Fixieren und Referenz-Codes"

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    ' Erst pruefen, DANN etwas veraendern - sonst bliebe der Plan bei
    ' einem Abbruch in einem halben Zustand zurueck.
    If n < 3 Then
        Notiz "Übersprungen - weniger als drei Inhaltszeilen."
        Exit Sub
    End If

    warFixiert = modWochenplan.IsPlanFixed()
    If warFixiert Then
        Notiz "Der Plan war fixiert. Für den Test wird die Fixierung aufgehoben " & _
              "und am Ende wieder gesetzt - die Referenz-Codes sind danach neu " & _
              "durchnummeriert (deshalb nur in der Testkopie!)."
        modSteuerung.Fixierung_Aufheben
    End If

    Chk "Vor dem Fixieren ist Spalte B eine Formel", ws.Cells(cr(1), "B").HasFormula

    modWochenplan.ClearLastError
    modSteuerung.Plan_Fixieren
    Chk "Fixieren ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()
    Chk "Der Plan gilt jetzt als fixiert", modWochenplan.IsPlanFixed()
    Chk "Spalte B ist jetzt ein fester Wert", Not ws.Cells(cr(1), "B").HasFormula

    ' Wandert der Code beim Verschieben mit?
    idx = 0
    For i = 2 To n
        If cr(i) - cr(i - 1) > 1 Then idx = i: Exit For
    Next i
    If idx = 0 Then idx = 3

    vor = Snap(ws)
    codeVor = T(ws.Cells(cr(idx), "B").Value)
    modWochenplan.ClearLastError
    modWochenplan.MoveBlock cr(idx), cr(idx), -1
    Chk "Verschieben im fixierten Zustand ohne Fehler", modWochenplan.LastError() = ""
    Chk "Es wurde tatsächlich etwas verschoben", Snap(ws) <> vor
    Chk "Der Referenz-Code ist mitgewandert", _
        T(ws.Cells(cr(idx - 1), "B").Value) = codeVor, _
        "erwartet " & codeVor & ", gefunden " & T(ws.Cells(cr(idx - 1), "B").Value)
    modWochenplan.MoveBlock cr(idx - 1), cr(idx - 1), 1
    Chk "Rundlauf im fixierten Zustand", Snap(ws) = vor

    ' Bekommt eine neue Zeile einen neuen, eindeutigen Code?
    naechste = modWochenplan.NextRefNumber(ws, lastRow)
    modWochenplan.InsertRowsBelow cr(2), cr(2)
    ws.Cells(cr(2) + 1, "F").Value = "TEST"
    ' Nicht AssignMissingCodes direkt aufrufen - Spalte B ist gesperrt.
    ' Der reguläre Weg hebt den Schutz auf und vergibt den Code dabei.
    modWochenplan.Wochenplan_Aktualisieren
    neuerCode = T(ws.Cells(cr(2) + 1, "B").Value)
    Chk "Neue Zeile bekommt den nächsten freien Code", _
        neuerCode = "TEST-" & Format$(naechste, "00"), _
        "erwartet TEST-" & Format$(naechste, "00") & ", gefunden " & neuerCode
    modWochenplan.DeleteRowsBlock cr(2) + 1, cr(2) + 1
    Chk "Die Testzeile ist wieder weg", modWochenplan.PlanLastRow(ws) = lastRow

    ' Ausgangszustand der Fixierung wiederherstellen
    If warFixiert Then
        Chk "Fixierung bleibt gesetzt (war vorher schon so)", modWochenplan.IsPlanFixed()
    Else
        modWochenplan.ClearLastError
        modSteuerung.Fixierung_Aufheben
        Chk "Fixierung aufheben ohne Fehlermeldung", modWochenplan.LastError() = ""
        Chk "Spalte B wird wieder berechnet", ws.Cells(cr(1), "B").HasFormula
    End If
End Sub


'=====================================================================
'  7  Blattschutz - und Arbeiten trotz aktivem Schutz
'=====================================================================
Private Sub T7_Blattschutz(ByVal ws As Worksheet)
    Dim st As Worksheet, sh As Worksheet
    Dim cr() As Long, n As Long, lastRow As Long
    Dim alleGeschuetzt As Boolean, r As Long, ferZeile As Long
    Dim i As Long, idx As Long, vor As String

    Abschnitt "7  Blattschutz"

    Set st = modWochenplan.SetSheet()

    modWochenplan.ClearLastError
    modSchutz.Blattschutz_Einrichten
    Chk "Blattschutz einrichten ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()

    alleGeschuetzt = True
    For Each sh In ThisWorkbook.Worksheets
        If Not sh.ProtectContents Then alleGeschuetzt = False
    Next sh
    Chk "Alle Blätter sind geschützt", alleGeschuetzt
    Chk "Die Mappenstruktur ist geschützt", ThisWorkbook.ProtectStructure

    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    If n < 2 Then
        Notiz "Zellprüfungen übersprungen - zu wenige Inhaltszeilen."
        Exit Sub
    End If

    Chk "Eingabezelle E ist entsperrt", ws.Cells(cr(1), "E").Locked = False
    Chk "Eingabezelle G ist entsperrt", ws.Cells(cr(1), "G").Locked = False
    Chk "Eingabezelle M ist entsperrt", ws.Cells(cr(1), "M").Locked = False
    Chk "Formelzelle C ist gesperrt", ws.Cells(cr(1), "C").Locked = True
    Chk "Hilfsspalte V ist gesperrt", ws.Cells(cr(1), "V").Locked = True

    ferZeile = 0
    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then ferZeile = r: Exit For
    Next r
    If ferZeile > 0 Then
        Chk "Ferienzeile " & ferZeile & " ist gesperrt", ws.Cells(ferZeile, "B").Locked = True
    End If

    If Not st Is Nothing Then
        If modWochenplan.StdCol() > 1 Then
            Chk "Stundenliste ist gesperrt", _
                st.Cells(modWochenplan.StdFirstRow(), modWochenplan.StdCol()).Locked = True
        End If
        Chk "Kalender ist bearbeitbar", _
            st.Cells(modWochenplan.WeekFirstRow(), "D").Locked = False
        Chk "Fixierungs-Schalter ist gesperrt", _
            st.Cells(modWochenplan.FixRow(), "B").Locked = True
    End If

    ' Das ist der Fall, an dem es zuletzt gescheitert ist:
    ' die Makros müssen bei AKTIVEM Blattschutz weiterarbeiten können.
    idx = 0
    For i = 2 To n
        If cr(i) - cr(i - 1) > 1 Then idx = i: Exit For
    Next i
    If idx = 0 Then idx = 2

    vor = Snap(ws)
    modWochenplan.ClearLastError
    modWochenplan.MoveBlock cr(idx), cr(idx), -1
    Chk "Verschieben bei aktivem Blattschutz ohne Fehlermeldung", _
        modWochenplan.LastError() = "", modWochenplan.LastError()
    ' Ohne diesen Nachweis wuerde der Rundlauf auch dann bestehen, wenn
    ' MoveBlock still ausgestiegen waere - genau das soll hier aber
    ' gerade geprueft werden.
    Chk "Es wurde bei aktivem Blattschutz tatsächlich verschoben", Snap(ws) <> vor
    modWochenplan.MoveBlock cr(idx - 1), cr(idx - 1), 1
    Chk "Rundlauf bei aktivem Blattschutz", Snap(ws) = vor
    Chk "Der Schutz ist danach wieder aktiv", modSchutz.SchutzAktiv()
End Sub


'=====================================================================
'  8  Titelblock und Seiteneinrichtung für das PDF
'=====================================================================
Private Sub T8_Kopf(ByVal ws As Worksheet)
    Dim lb As Worksheet, st As Worksheet
    Dim kopfDa As Boolean
    Dim kopfVorher As Long, kopfNachher As Long, kopfZweimal As Long
    Dim vor As String, nach As String
    Dim titel As String, logoDa As Boolean
    Dim bereich As String, ausrichtung As Long, wiederhol As String
    Dim gelesen As Boolean

    Abschnitt "8  Titelblock und PDF"

    Set st = modWochenplan.SetSheet()
    Set lb = Blatt(LB_SHEET)

    ' Kopf_Aktualisieren steigt im leisen Modus still aus, wenn die
    ' Kopfangaben fehlen - das muss vorher geprüft werden, sonst
    ' bestünde die Fehlerprüfung unten immer.
    kopfDa = (modWochenplan.KopfZeileFach() > 0)
    Chk "Kopfangabe """ & LBL_FACH & """ in den Einstellungen vorhanden", kopfDa
    Chk "Kopfangabe """ & LBL_KLASSE & """ vorhanden", modWochenplan.KopfZeileKlasse() > 0
    Chk "Kopfangabe """ & LBL_SCHULE & """ vorhanden", modWochenplan.KopfZeileSchule() > 0
    Chk "Kopfangabe """ & LBL_LEHRER & """ vorhanden", modWochenplan.KopfZeileLehrer() > 0
    If Not kopfDa Then Exit Sub

    ' --- Titelblock schreiben ---------------------------------------
    vor = Snap(ws)
    kopfVorher = modWochenplan.WP_HEAD_ROW()

    modWochenplan.ClearLastError
    modKopf.Kopf_Aktualisieren True
    Chk "Titelblock schreiben ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()

    modWochenplan.ResetKopfzeilen
    kopfNachher = modWochenplan.WP_HEAD_ROW()

    ' Der Block sitzt über der Überschriftenzeile - die muss also
    ' mindestens um seine Höhe nach unten gewandert sein bzw. bereits
    ' unterhalb davon liegen.
    ' --- Der Plan selbst darf sich dabei NICHT verändert haben -----
    ' Snap liest über PlanFirstRow/PlanLastRow, also relativ zur
    ' Kopfzeile. Wenn der Titelblock die Daten verschoben statt sie
    ' nur nach unten geschoben hätte, fällt es hier auf.
    nach = Snap(ws)
    Chk "Der Plan ist durch den Titelblock unverändert", nach = vor

    ' --- Zweiter Lauf darf nichts doppeln ----------------------------
    ' Der häufigste Fehler bei so einem Block: bei jedem Aufruf
    ' schiebt er vier weitere Zeilen ein.
    modWochenplan.ClearLastError
    modKopf.Kopf_Aktualisieren True
    modWochenplan.ResetKopfzeilen
    kopfZweimal = modWochenplan.WP_HEAD_ROW()
    Chk "Zweiter Aufruf fügt keinen zweiten Block ein", _
        kopfZweimal = kopfNachher, _
        "nach 1x: " & kopfNachher & ", nach 2x: " & kopfZweimal
    Chk "Der Plan ist auch nach dem zweiten Lauf unverändert", Snap(ws) = vor

    ' --- Inhalt des Blocks ------------------------------------------
    ' Die Titelzeile ist in Zonen zerlegt (links Logo, Mitte Fach und
    ' Klasse, rechts Lehrkraft und Schuljahr). Deshalb NICHT eine
    ' feste Zelle abfragen, sondern die ganze Zeile einsammeln.
    Chk "Der Titelblock ist genau " & modKopf.KOPF_ZEILEN & " Zeilen hoch", _
        kopfNachher = modKopf.KOPF_ZEILEN + 1, "Kopfzeile " & kopfNachher

    titel = ZeilenText(ws, 1, modKopf.PDF_WP_SPALTE_ERSTE, modKopf.PDF_WP_SPALTE_LETZTE)
    Chk "Titelzeile im " & WP_SHEET & " ist gefüllt", Len(titel) > 0, "'" & titel & "'"
    Chk "Fach steht im Titel des " & WP_SHEET, _
        InStr(1, titel, KopfWertTest(st, modWochenplan.KopfZeileFach()), vbTextCompare) > 0, _
        "'" & titel & "'"
    Chk "Die Titelzeile ist in Zonen verbunden", ws.Cells(1, "B").MergeCells
    If Not lb Is Nothing Then
        Chk "Titelzeile in " & LB_SHEET & " ist gefüllt", _
            Len(ZeilenText(lb, 1, modKopf.PDF_LB_SPALTE_ERSTE, _
                           modKopf.PDF_LB_SPALTE_LETZTE)) > 0
    End If
    Notiz "Titelzeile: '" & titel & "'"

    logoDa = modKopf.LogoVorhanden(ws)
    Chk "Ein Logo ist sichtbar", logoDa
    If logoDa Then
        ' Das Seitenverhaeltnis ist die eigentliche Probe: am
        ' 03.09.2026 war das Logo auf 66x55 zusammengedrueckt und
        ' voellig unlesbar. Die Schullogos sind breite Banner.
        Chk "Das Logo ist nicht verzerrt (Breite deutlich groesser als Hoehe)", _
            LogoVerhaeltnis(ws) > 3, "Verhaeltnis " & Format$(LogoVerhaeltnis(ws), "0.00")
        Notiz "Logo-Seitenverhaeltnis: " & Format$(LogoVerhaeltnis(ws), "0.00")
    End If

    ' Der farbige Rahmen um die Titelzeile
    Chk "Die Titelzeile hat einen Rahmen", _
        ws.Cells(1, "B").Borders(xlEdgeTop).LineStyle <> xlNone

    ' --- Seiteneinrichtung fürs PDF --------------------------------
    ' Nur einrichten, NICHT exportieren: der Export würde eine Datei
    ' hinterlassen und kann am fehlenden Druckertreiber scheitern.
    modWochenplan.ClearLastError
    modKopf.SeiteFuerPdfEinrichten ws, xlPaperA4
    Chk "Seiteneinrichtung ohne Fehlermeldung", modWochenplan.LastError() = "", _
        modWochenplan.LastError()

    gelesen = False
    On Error Resume Next
    Err.Clear
    bereich = ws.PageSetup.PrintArea
    ausrichtung = ws.PageSetup.Orientation
    wiederhol = ws.PageSetup.PrintTitleRows
    gelesen = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Chk "Seiteneinrichtung ist auslesbar", gelesen
    If gelesen Then
        Chk "Druckbereich beginnt in Spalte " & modKopf.PDF_WP_SPALTE_ERSTE & " Zeile 1", _
            InStr(1, bereich, "$" & modKopf.PDF_WP_SPALTE_ERSTE & "$1:", vbTextCompare) > 0, _
            "'" & bereich & "'"
        Chk "Druckbereich endet in Spalte " & modKopf.PDF_WP_SPALTE_LETZTE, _
            InStr(1, bereich, "$" & modKopf.PDF_WP_SPALTE_LETZTE & "$", vbTextCompare) > 0, _
            "'" & bereich & "'"
        Chk "Querformat ist eingestellt", ausrichtung = xlLandscape
        Chk "Die Überschriftenzeile wird auf jeder Seite wiederholt", _
            InStr(1, wiederhol, "$" & kopfNachher & ":", vbTextCompare) > 0, _
            "'" & wiederhol & "'"
        Notiz "Druckbereich: '" & bereich & "'"
    End If

    ' --- Gestrichelte Seitenumbruchlinien ----------------------------
    modWochenplan.SeitenumbruecheAus
    Chk "Die gestrichelten Seitenumbruchlinien sind aus", _
        ws.DisplayPageBreaks = False
End Sub


'=====================================================================
'  9  Fokus: liegt man nach jeder Schaltflaeche auf dem richtigen Blatt?
'=====================================================================
'  Der entscheidende Kniff: vor jedem Aufruf wird BEWUSST ein anderes
'  Blatt aktiviert. Startet man auf dem Blatt, auf dem man am Ende sein
'  soll, besteht die Pruefung immer - und genau deshalb ist der Sprung
'  auf die Anleitung bei "Blattschutz aus" zweimal durchgerutscht.
'
'  Dieser Abschnitt laeuft ZULETZT: Ferien_Aus_Kalender_Vorschlagen
'  schreibt die Ferientabelle neu, das darf die frueheren Abschnitte
'  nicht stoeren.
'=====================================================================
Private Sub T9_Fokus(ByVal ws As Worksheet)
    Dim warFixiert As Boolean, warGeschuetzt As Boolean

    Abschnitt "9  Fokus nach den Schaltflaechen"

    warFixiert = modWochenplan.IsPlanFixed()
    warGeschuetzt = modSchutz.SchutzAktiv()

    ' --- Blatt "Steuerung" -------------------------------------------
    FokusProbe "Wochenplan neu aufbauen", WP_SHEET, 1
    FokusProbe "Stoffverteilungsplan fixieren", CTRL_SHEET, 2
    FokusProbe "Fixierung aufheben", CTRL_SHEET, 3
    FokusProbe "Blattschutz ein / aus (ausschalten)", CTRL_SHEET, 4
    FokusProbe "Blattschutz ein / aus (einschalten)", CTRL_SHEET, 5

    ' --- Blatt "Einstellungen" ---------------------------------------
    FokusProbe "Kopf übernehmen", SET_SHEET, 6
    FokusProbe "Ferien aus Kalender vorschlagen", SET_SHEET, 7

    ' --- Ausgangszustand wiederherstellen ----------------------------
    If warFixiert And Not modWochenplan.IsPlanFixed() Then modSteuerung.Plan_Fixieren
    If Not warFixiert And modWochenplan.IsPlanFixed() Then modSteuerung.Fixierung_Aufheben
    If warGeschuetzt And Not modSchutz.SchutzAktiv() Then modSchutz.Blattschutz_Einrichten
    Chk "Fixierung ist wie vor Abschnitt 9", modWochenplan.IsPlanFixed() = warFixiert
    Chk "Blattschutz ist wie vor Abschnitt 9", modSchutz.SchutzAktiv() = warGeschuetzt
End Sub


'=====================================================================
'  10  Uebernahme aus einer bisherigen Datei
'  ------------------------------------------------------------------
'  Der Abschnitt spielt einen echten Umstieg durch:
'
'    1. Diese Mappe wird als temporaere Datei weggeschrieben - das ist
'       die "bisherige Datei".
'    2. Was drinsteht, wird gemerkt.
'    3. Wochenplan und Lernbereiche dieser Mappe werden LEERGERAEUMT.
'    4. modUebernahme holt die Daten aus der temporaeren Datei zurueck.
'    5. Verglichen wird Zelle fuer Zelle.
'
'  Nur so faellt ein Fehler auf der SCHREIBSEITE auf. Die erste
'  Fassung von modUebernahme hat die Quelle vollstaendig gelesen und
'  trotzdem nur eine einzige Zeile geschrieben - ein Test, der bloss
'  die Lesefunktionen prueft, haette das durchgewunken.
'
'  Der Abschnitt laeuft ALS LETZTER: er raeumt die Mappe zwischendrin
'  leer. Geht dabei etwas schief, bleibt die temporaere Datei liegen
'  und ihr Pfad steht im Bericht - es ist also nichts verloren.
'=====================================================================
Private Sub T10_Uebernahme(ByVal ws As Worksheet)
    Dim tmp As String, altName As String
    Dim nPlan As Long, nLb As Long, nWo As Long, nFer As Long
    Dim sollPlan() As String, sollLb() As String
    Dim nSollPlan As Long, nSollLb As Long
    Dim cr() As Long, n As Long, i As Long
    Dim lastRow As Long, abw As Long, ersteAbw As String
    Dim zl As Worksheet, zSum As Long

    Abschnitt "10  Uebernahme aus einer bisherigen Datei"

    Set zl = modWochenplan.LbSheet()
    If zl Is Nothing Then
        Chk "Blatt " & LB_SHEET & " vorhanden", False
        Exit Sub
    End If

    ' --- 1. Sollzustand merken ---------------------------------------
    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    If n < 2 Then
        Notiz "Zu wenige Planzeilen fuer diesen Abschnitt - uebersprungen."
        Exit Sub
    End If
    ReDim sollPlan(1 To n)
    For i = 1 To n
        sollPlan(i) = PlanZeileText(ws, cr(i))
    Next i
    nSollPlan = n

    zSum = modWochenplan.LbSummeRow(zl)
    nSollLb = modWochenplan.LbLastDataRow(zl, zSum) - modWochenplan.LB_FIRST_ROW() + 1
    If nSollLb < 1 Then nSollLb = 0
    If nSollLb > 0 Then
        ReDim sollLb(1 To nSollLb)
        For i = 1 To nSollLb
            sollLb(i) = LbZeileText(zl, modWochenplan.LB_FIRST_ROW() + i - 1)
        Next i
    End If

    ' --- 2. Kopie anlegen --------------------------------------------
    tmp = TempPfad("Uebernahme_Quelle.xlsm")
    On Error Resume Next
    Kill tmp
    Err.Clear
    ThisWorkbook.SaveCopyAs tmp
    If Err.Number <> 0 Then
        Chk "Kopie als Quelle anlegen", False, "Fehler " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Chk "Kopie als Quelle angelegt", (Len(Dir$(tmp)) > 0), tmp

    ' --- 3. Mappe leerraeumen ----------------------------------------
    MappeLeeren ws, zl
    Chk "Mappe gilt danach als leer", modUebernahme.ZielIstLeer()
    LeerBegriffPruefen ws

    ' --- 4. Uebernahme -----------------------------------------------
    nPlan = modUebernahme.UebernahmeAusfuehren(tmp, altName, nLb, nWo, nFer)
    Chk "Uebernahme ohne Fehler", (nPlan >= 0), "Rueckgabe " & nPlan
    If nPlan < 0 Then Exit Sub

    '  DIESELBEN Post-Schritte wie hinter der Schaltflaeche. Ohne sie
    '  prueft der Abschnitt einen Ablauf, den es gar nicht gibt:
    '  UebernahmeAusfuehren schreibt nur die Daten, alles Berechnete
    '  baut erst Setup auf. Beim ersten Lauf ist genau daran der Check
    '  "Rechenspalte F ist wieder eine Formel" gescheitert - zu Recht,
    '  denn F war leer. Der Fehler sass im Test, nicht im Programm.
    modSteuerung.Setup_Stoffverteilungsplan
    modKalender.UW_Und_Ferien_Generieren

    Chk "alle " & nSollPlan & " Planzeilen uebernommen", (nPlan = nSollPlan), _
        "uebernommen: " & nPlan
    Chk "alle " & nSollLb & " Lernbereiche uebernommen", (nLb = nSollLb), _
        "uebernommen: " & nLb
    Chk "Schulwochen-Kalender uebernommen", _
        (nWo = modWochenplan.WeekLastRow() - modWochenplan.WeekFirstRow() + 1), _
        "uebernommen: " & nWo

    ' --- 5. Zelle fuer Zelle vergleichen ------------------------------
    lastRow = modWochenplan.PlanLastRow(ws)
    n = modWochenplan.ContentRows(ws, lastRow, cr)
    abw = 0: ersteAbw = ""
    For i = 1 To nSollPlan
        If i > n Then
            abw = abw + 1
            If ersteAbw = "" Then ersteAbw = "Zeile " & i & " fehlt ganz"
        ElseIf PlanZeileText(ws, cr(i)) <> sollPlan(i) Then
            abw = abw + 1
            If ersteAbw = "" Then
                ersteAbw = "Zeile " & i & " (Blattzeile " & cr(i) & "): erwartet [" & _
                           Left$(sollPlan(i), 60) & "] / gefunden [" & _
                           Left$(PlanZeileText(ws, cr(i)), 60) & "]"
            End If
        End If
    Next i
    Chk "Wochenplan Zelle fuer Zelle gleich", (abw = 0), _
        abw & " Abweichung(en); " & ersteAbw

    abw = 0: ersteAbw = ""
    For i = 1 To nSollLb
        If LbZeileText(zl, modWochenplan.LB_FIRST_ROW() + i - 1) <> sollLb(i) Then
            abw = abw + 1
            If ersteAbw = "" Then ersteAbw = "Lernbereich " & i
        End If
    Next i
    Chk "Lernbereiche Zelle fuer Zelle gleich", (abw = 0), _
        abw & " Abweichung(en); " & ersteAbw

    ' --- 6. Die Summenzeile muss die Datenzeilen treffen --------------
    zSum = modWochenplan.LbSummeRow(zl)
    Chk "Summenzeile steht direkt unter dem letzten Lernbereich", _
        (zSum = modWochenplan.LB_FIRST_ROW() + nSollLb), _
        "Summenzeile " & zSum & ", erwartet " & (modWochenplan.LB_FIRST_ROW() + nSollLb)
    If zSum > 0 Then
        Chk "Summenformel zeigt auf die Datenzeilen", _
            (InStr(1, zl.Cells(zSum, "E").Formula, _
                   "E" & modWochenplan.LB_FIRST_ROW() & ":E" & (zSum - 1)) > 0), _
            zl.Cells(zSum, "E").Formula
    End If
    Chk "Rechenspalte F ist wieder eine Formel", _
        (Left$(zl.Cells(modWochenplan.LB_FIRST_ROW(), "F").Formula, 1) = "="), _
        zl.Cells(modWochenplan.LB_FIRST_ROW(), "F").Formula

    ' --- 7. Zweiter Anlauf muss abgelehnt werden ----------------------
    Chk "volle Mappe wird nicht mehr als leer angesehen", Not modUebernahme.ZielIstLeer()

    ' --- 8. Aufraeumen ------------------------------------------------
    On Error Resume Next
    Kill tmp
    On Error GoTo 0
    Chk "temporaere Quelldatei aufgeraeumt", (Len(Dir$(tmp)) = 0)
End Sub


'  Was macht eine Mappe "beschrieben"? Wer die Uebernahme in eine als
'  leer erkannte Mappe laufen laesst, ueberschreibt sie ohne Rueckfrage.
'  Entschieden ist (Issue #7): nur F, G, H, I und M zaehlen.
'
'  Die Spaltenliste steht hier ein zweites Mal - als Kopie, nicht als
'  Gegenprobe. Sie zeigt nur, wenn eine der fuenf Spalten aus
'  MappeIstLeer herausfaellt. Das eigentliche Gewicht liegt auf der
'  zweiten Schleife, den Spalten, die NICHT zaehlen duerfen:
'    E schreibt UW_Und_Ferien_Generieren maschinell in jede Zeile,
'    J setzt ZeilenEinfuegen in jeder Zeile auf False.
'  Kaeme eine der beiden dazu, waere keine Mappe je wieder leer und die
'  Uebernahme dauerhaft unbenutzbar. K ist der Fall, ueber den der
'  Nutzer entschieden hat.
'
'  Setzt voraus, dass die Mappe gerade leergeraeumt wurde.
Private Sub LeerBegriffPruefen(ByVal ws As Worksheet)
    Dim sp As Variant, r As Long

    '  Kein stiller Ausstieg: eine uebersprungene Probe sieht im
    '  Bericht sonst genauso aus wie eine bestandene.
    r = modWochenplan.WP_FIRST_ROW()
    Chk "Spaltenprobe laeuft: erste Planzeile ist keine Ferienzeile", _
        Not modWochenplan.IsFerienRow(ws, r)
    If modWochenplan.IsFerienRow(ws, r) Then Exit Sub

    modWochenplan.FastOn ws
    For Each sp In Array("F", "G", "H", "I", "M")
        ws.Cells(r, CStr(sp)).Value = "x"
        Chk "nur Spalte " & sp & " gefuellt: Mappe gilt als beschrieben", _
            Not modUebernahme.ZielIstLeer()
        ws.Cells(r, CStr(sp)).ClearContents
    Next sp

    For Each sp In Array("E", "J", "K")
        ws.Cells(r, CStr(sp)).Value = "x"
        '  Ohne diesen Zwischenschritt waere die naechste Pruefung auch
        '  dann gruen, wenn das Schreiben gar nicht angekommen ist -
        '  sie erwartet ja "leer".
        Chk "Spalte " & sp & " liess sich ueberhaupt beschreiben", _
            (Len(Trim$(CStr(ws.Cells(r, CStr(sp)).Value))) > 0)
        Chk "nur Spalte " & sp & " gefuellt: Mappe gilt weiter als leer", _
            modUebernahme.ZielIstLeer()
        ws.Cells(r, CStr(sp)).ClearContents
    Next sp
    modWochenplan.FastOff

    Chk "nach der Spaltenprobe ist die Mappe wieder leer", _
        modUebernahme.ZielIstLeer()
End Sub


'  Der Inhalt einer Planzeile in genau den Spalten, die uebernommen
'  werden. Alles andere (A bis D, L, Hilfsspalten) sind Formeln und
'  werden neu gebaut - die duerfen sich unterscheiden.
Private Function PlanZeileText(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Variant, s As String
    For Each c In Array("E", "F", "G", "H", "I", "J", "K", "M")
        s = s & "|" & Trim$(CStr(ws.Cells(r, CStr(c)).Value))
    Next c
    PlanZeileText = s
End Function


Private Function LbZeileText(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Variant, s As String
    For Each c In Array(1, 2, 3, 4, 5, 10)
        s = s & "|" & Trim$(CStr(ws.Cells(r, CLng(c)).Value))
    Next c
    LbZeileText = s
End Function


'  Wochenplan und Lernbereiche in den Zustand einer frischen Vorlage
'  bringen: eine leere Planzeile, keine Lernbereiche. Der Blattschutz
'  muss dafuer aus sein - deshalb FastOn/FastOff drumherum.
Private Sub MappeLeeren(ByVal ws As Worksheet, ByVal zl As Worksheet)
    Dim lastRow As Long, erste As Long, zSum As Long

    modWochenplan.FastOn
    erste = modWochenplan.WP_FIRST_ROW()
    lastRow = modWochenplan.PlanLastRow(ws)
    If lastRow > erste Then ws.Rows(erste + 1 & ":" & lastRow).Delete Shift:=xlUp
    ws.Range(ws.Cells(erste, "E"), ws.Cells(erste, "M")).ClearContents

    zSum = modWochenplan.LbSummeRow(zl)
    If zSum > modWochenplan.LB_FIRST_ROW() Then
        zl.Range(zl.Cells(modWochenplan.LB_FIRST_ROW(), "A"), _
                 zl.Cells(zSum - 1, "J")).ClearContents
    End If
    modWochenplan.ResetKopfzeilen
    modWochenplan.FastOff
End Sub


'  Bewusst IMMER der Temp-Ordner und nicht der Mappenordner: bei
'  eingeschaltetem AutoSpeichern ist ThisWorkbook.Path eine
'  SharePoint-Adresse, und selbst der lokale OneDrive-Ordner ist fuer
'  eine Datei, die Sekunden spaeter wieder verschwindet, der falsche
'  Platz - die Synchronisierung laedt sie sonst noch hoch.
Private Function TempPfad(ByVal name As String) As String
    Dim p As String
    p = Environ$("TEMP")
    If p = "" Then p = Environ$("USERPROFILE")
    If p = "" Then p = ThisWorkbook.Path
    TempPfad = p & "\~Selbsttest_" & name
End Function



'=====================================================================
'  D E R   T E S T   F U E R   D E N   T E S T
'  ------------------------------------------------------------------
'  Ein gruener Selbsttest heisst nur dann etwas, wenn seine Pruefungen
'  bei einem echten Fehler auch wirklich rot werden. Dass das nicht
'  selbstverstaendlich ist, hat dieses Projekt teuer gelernt:
'
'    Der Check "Keine leeren Zeilen am Ende des Plans" hat "ok"
'    gemeldet, waehrend 35 leere Zeilen unter der Tabelle standen.
'    Er begann bei PlanLastRow und suchte nach OBEN - die Zeilen lagen
'    darunter. Der Check hatte denselben blinden Fleck wie der Fehler,
'    den er finden sollte.
'
'  Selbsttest_Pruefen baut deshalb definierte Fehler ein und verlangt,
'  dass GENAU der dafuer zustaendige Check rot wird - und dass er nach
'  dem Zuruecknehmen wieder gruen ist. Was hier nicht anschlaegt, ist
'  kein Test, sondern Dekoration.
'
'  Angesetzt wird nur an den Abschnitten 0 und 1b: die fassen nichts
'  an und reparieren nichts, eine Sabotage bleibt also bis zur Pruefung
'  stehen. In einem Abschnitt, der selbst aufraeumt, waere das Ergebnis
'  wertlos.
'=====================================================================
Public Sub Selbsttest_Pruefen()
    Dim ws As Worksheet, i As Long, pfad As String
    Dim vorher As Long, rot As Boolean, wiederGruen As Boolean
    Dim erkannt As Long, offenGeblieben As Long, t0 As Single

    If InStr(1, ThisWorkbook.Name, "test", vbTextCompare) = 0 Then
        MsgBox "Der Mutationstest laeuft nur in einer Kopie." & vbCrLf & vbCrLf & _
               "Er baut absichtlich Fehler in die Mappe ein und nimmt sie wieder " & _
               "zurueck. Bitte in einer Datei starten, deren Name ""Test"" enthaelt.", _
               vbOKOnly, "Selbsttest pruefen"
        Exit Sub
    End If

    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Sub

    mRep = "": mOk = 0: mFail = 0: mFehlerListe = ""
    t0 = Timer
    Zeile "=========================================================="
    Zeile " Mutationstest - schlaegt der Selbsttest bei Fehlern an?"
    Zeile " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Zeile " Datei  : " & ThisWorkbook.Name
    Zeile "=========================================================="

    modWochenplan.SetQuiet True
    On Error GoTo Fail

    ' --- Ausgangszustand herstellen und pruefen ----------------------
    ' Ohne gesunde Mappe sagt der ganze Lauf nichts aus: ein Check, der
    ' schon vorher rot ist, wird durch jede Sabotage "erkannt".
    Abschnitt "Ausgangszustand"
    modSteuerung.Setup_Stoffverteilungsplan
    modKalender.UW_Und_Ferien_Generieren
    vorher = mFail
    AbschnittLaufen 0, ws
    AbschnittLaufen 1, ws
    Chk "Die Mappe ist vor dem ersten Eingriff fehlerfrei", (mFail = vorher), _
        (mFail - vorher) & " Pruefung(en) schon ohne Sabotage rot"
    If mFail > vorher + 1 Then
        Zeile ""
        Zeile "*** ABBRUCH: erst die Mappe in Ordnung bringen."
        GoTo Ende
    End If

    ' --- Die Mutationen ----------------------------------------------
    For i = 1 To MUT_ANZAHL
        Abschnitt "Mutation " & i & "  " & MutName(i)
        Notiz "erwartet rot: """ & MutErwartet(i) & """ in Abschnitt " & MutAbschnitt(i)

        Sabotieren i, ws
        mErwarte = MutErwartet(i)
        mGetroffen = False
        vorher = mFail
        AbschnittLaufen MutAbschnitt(i), ws
        rot = mGetroffen
        mErwarte = ""

        Zuruecknehmen i, ws
        vorher = mFail
        AbschnittLaufen MutAbschnitt(i), ws
        wiederGruen = (mFail = vorher)

        ' Die beiden Laeufe oben zaehlen als Pruefungen mit - das
        ' verfaelscht die Bilanz. Deshalb wird hier ausdruecklich nur
        ' das Ergebnis der Mutation bewertet.
        Chk "Mutation " & i & " wird erkannt: " & MutName(i), rot, _
            "der Check """ & MutErwartet(i) & """ ist NICHT rot geworden"
        Chk "Mutation " & i & " ist sauber zurueckgenommen", wiederGruen
        If rot Then erkannt = erkannt + 1
        If Not wiederGruen Then offenGeblieben = offenGeblieben + 1
    Next i

Ende:
    modWochenplan.SetQuiet False
    Zeile ""
    Zeile "=========================================================="
    Zeile " " & erkannt & " von " & MUT_ANZAHL & " Mutationen wurden erkannt"
    If offenGeblieben > 0 Then
        Zeile " ACHTUNG: " & offenGeblieben & " Sabotage(n) wurden NICHT sauber"
        Zeile " zurueckgenommen. Die Mappe ist nicht im Ausgangszustand."
    End If
    Zeile " Dauer: " & Format$(Timer - t0, "0.0") & " s"
    Zeile "=========================================================="
    pfad = BerichtSchreiben("Selbsttest_Mutationen.txt")

    MsgBox IIf(erkannt = MUT_ANZAHL, "Alle Mutationen wurden erkannt.", _
                                     "ACHTUNG: " & (MUT_ANZAHL - erkannt) & _
                                     " Mutation(en) sind dem Selbsttest ENTGANGEN.") & _
           vbCrLf & vbCrLf & erkannt & " von " & MUT_ANZAHL & " erkannt." & _
           IIf(offenGeblieben = 0, "", vbCrLf & vbCrLf & _
               "ACHTUNG: " & offenGeblieben & " Sabotage(n) wurden nicht sauber " & _
               "zurueckgenommen. Die Mappe ist NICHT im Ausgangszustand - bitte " & _
               "einmal ""Einrichtung / Reparatur"" druecken.") & _
           IIf(mFail = 0, "", vbCrLf & vbCrLf & "Auffaelligkeiten:" & mFehlerListe) & _
           vbCrLf & vbCrLf & "Bericht:" & vbCrLf & pfad, vbOKOnly, "Selbsttest pruefen"
    Exit Sub

Fail:
    modWochenplan.SetQuiet False
    modWochenplan.FastOff
    Zeile ""
    Zeile "*** ABBRUCH bei Mutation " & i & ": Fehler " & Err.Number & " - " & Err.Description
    Zeile "*** Die Mappe kann veraendert zurueckbleiben - bitte einmal"
    Zeile "*** ""Einrichtung / Reparatur"" druecken."
    pfad = BerichtSchreiben("Selbsttest_Mutationen.txt")
    MsgBox "Der Mutationstest ist abgebrochen." & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Mutation: " & i & vbCrLf & vbCrLf & _
           "Die Mappe kann veraendert zurueckbleiben - bitte einmal " & _
           """Einrichtung / Reparatur"" druecken." & vbCrLf & vbCrLf & _
           "Bericht:" & vbCrLf & pfad, vbOKOnly, "Selbsttest pruefen"
End Sub


Private Sub AbschnittLaufen(ByVal nr As Long, ByVal ws As Worksheet)
    Select Case nr
        Case 0: T0_Erreichbarkeit ws
        Case 1: T1b_Zustand ws
    End Select
End Sub


Private Function MutName(ByVal i As Long) As String
    Select Case i
        Case 1: MutName = "Leerzeile unter der Tabelle"
        Case 2: MutName = "Rechenspalte F der Lernbereiche geleert"
        Case 3: MutName = "Summenformel auf einen falschen Bereich gestellt"
        Case 4: MutName = "Kennzeichen einer Ferienzeile entfernt"
        Case 5: MutName = "Blatt ""Update"" wieder angelegt"
        Case 6: MutName = "Stand der Anleitung verstellt"
        Case 7: MutName = "Blattschutz aufgehoben"
        Case 8: MutName = "Unterrichtswochen verdreht"
        Case 9: MutName = "Makro von einer Schaltflaeche entfernt"
        Case 10: MutName = "zwei Schaltflaechen mit demselben Makro"
    End Select
End Function


Private Function MutErwartet(ByVal i As Long) As String
    Select Case i
        Case 1: MutErwartet = "formatierten Leerzeilen"
        Case 2: MutErwartet = "Rechenspalten F bis I"
        Case 3: MutErwartet = "Summenformel deckt"
        Case 4: MutErwartet = "Kennzeichen in Spalte"
        Case 5: MutErwartet = "Kein Update-Blatt"
        Case 6: MutErwartet = "aktuellen Stand"
        Case 7: MutErwartet = "Blattschutz ist aktiv"
        Case 8: MutErwartet = "rueckwaerts"
        Case 9: MutErwartet = "zeigen auf ein Makro"
        Case 10: MutErwartet = "eigenes Makro"
    End Select
End Function


Private Function MutAbschnitt(ByVal i As Long) As Long
    If i >= 9 Then MutAbschnitt = 0 Else MutAbschnitt = 1
End Function


'  Sabotage und Ruecknahme stehen bewusst direkt untereinander -
'  so faellt sofort auf, wenn eine Ruecknahme nicht zur Sabotage passt.
Private Sub Sabotieren(ByVal i As Long, ByVal ws As Worksheet)
    Dim zl As Worksheet, r As Long, z1 As Long, zSum As Long
    Dim neu As Worksheet, s As Shape, s1 As Shape

    Set zl = modWochenplan.LbSheet()
    Select Case i

        Case 1      ' eine Zeile unter der Tabelle auf Planzeilen-Hoehe
            r = modWochenplan.PlanLastRow(ws) + 1
            modWochenplan.FastOn
            mSaveZahl = ws.Rows(r).RowHeight
            ws.Rows(r).RowHeight = ws.Rows(r - 1).RowHeight
            modWochenplan.FastOff

        Case 2      ' Rechenspalte F leeren
            z1 = modWochenplan.LB_FIRST_ROW()
            modWochenplan.FastOn
            mSaveText = zl.Cells(z1, "F").Formula
            zl.Cells(z1, "F").ClearContents
            modWochenplan.FastOff

        Case 3      ' Summenformel auf einen falschen Bereich
            zSum = modWochenplan.LbSummeRow(zl)
            modWochenplan.FastOn
            mSaveText = zl.Cells(zSum, "E").Formula
            zl.Cells(zSum, "E").Formula = "=SUM(E1:E1)"
            modWochenplan.FastOff

        Case 4      ' Kennzeichen der ersten Ferienzeile entfernen
            modWochenplan.FastOn
            mSaveZahl = 0
            For r = WP_FIRST_ROW To modWochenplan.PlanLastRow(ws)
                If modWochenplan.IsFerienRow(ws, r) Then
                    mSaveZahl = r
                    mSaveText = CStr(ws.Cells(r, MARK_COL).Value)
                    ws.Cells(r, MARK_COL).ClearContents
                    Exit For
                End If
            Next r
            modWochenplan.FastOff

        Case 5      ' Blatt "Update" anlegen
            On Error Resume Next
            Application.DisplayAlerts = False
            ThisWorkbook.Unprotect SCHUTZ_PW
            Set neu = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            neu.Name = "Update"
            ThisWorkbook.Protect Password:=SCHUTZ_PW, Structure:=True, Windows:=False
            Application.DisplayAlerts = True
            On Error GoTo 0

        Case 6      ' Stand der Anleitung verstellen
            On Error Resume Next
            ThisWorkbook.Names(modAnleitung.STAND_NAME).Delete
            ThisWorkbook.Names.Add Name:=modAnleitung.STAND_NAME, _
                                   RefersTo:="=""veraltet""", Visible:=False
            On Error GoTo 0

        Case 7      ' Blattschutz aufheben
            If modSchutz.SchutzAktiv() Then modSchutz.Blattschutz_Umschalten

        Case 8      ' zwei Unterrichtswochen vertauschen
            modWochenplan.FastOn
            mSaveZahl = 0
            For r = modWochenplan.PlanLastRow(ws) To WP_FIRST_ROW + 1 Step -1
                If Not modWochenplan.IsFerienRow(ws, r) Then
                    mSaveZahl = r
                    mSaveText = CStr(ws.Cells(r, "E").Value)
                    ws.Cells(r, "E").Value = 1
                    Exit For
                End If
            Next r
            modWochenplan.FastOff

        Case 9      ' einer Schaltflaeche das Makro nehmen
            Set s = SteuerungsForm(1)
            If Not s Is Nothing Then
                mSaveName = s.Name
                mSaveText = s.OnAction
                s.OnAction = ""
            End If

        Case 10     ' zwei Schaltflaechen auf dasselbe Makro zeigen lassen
            '  Der stille Zwilling: keine Schaltflaeche ist leer, aber
            '  eine tut das Falsche. Genau dieser Zustand ist am
            '  05.09.2026 durch die kaputte Ruecknahme von Mutation 9
            '  entstanden - und der damalige Check hat ihn NICHT
            '  gesehen, weil er nur auf leere Ziele geachtet hat.
            Set s = SteuerungsForm(2)
            Set s1 = SteuerungsForm(1)
            If Not s Is Nothing And Not s1 Is Nothing Then
                mSaveName = s.Name
                mSaveText = s.OnAction
                s.OnAction = s1.OnAction
            End If

    End Select
End Sub


Private Sub Zuruecknehmen(ByVal i As Long, ByVal ws As Worksheet)
    Dim zl As Worksheet, r As Long, z1 As Long, zSum As Long, s As Shape

    Set zl = modWochenplan.LbSheet()
    Select Case i

        Case 1
            r = modWochenplan.PlanLastRow(ws) + 1
            modWochenplan.FastOn
            ws.Rows(r).RowHeight = mSaveZahl
            modWochenplan.FastOff

        Case 2
            z1 = modWochenplan.LB_FIRST_ROW()
            modWochenplan.FastOn
            zl.Cells(z1, "F").Formula = mSaveText
            modWochenplan.FastOff

        Case 3
            zSum = modWochenplan.LbSummeRow(zl)
            modWochenplan.FastOn
            zl.Cells(zSum, "E").Formula = mSaveText
            modWochenplan.FastOff

        Case 4
            If mSaveZahl > 0 Then
                modWochenplan.FastOn
                modWochenplan.MarkierungSetzen ws, CLng(mSaveZahl)
                modWochenplan.FastOff
            End If

        Case 5
            On Error Resume Next
            Application.DisplayAlerts = False
            ThisWorkbook.Unprotect SCHUTZ_PW
            ThisWorkbook.Worksheets("Update").Delete
            ThisWorkbook.Protect Password:=SCHUTZ_PW, Structure:=True, Windows:=False
            Application.DisplayAlerts = True
            On Error GoTo 0

        Case 6
            On Error Resume Next
            ThisWorkbook.Names(modAnleitung.STAND_NAME).Delete
            ThisWorkbook.Names.Add Name:=modAnleitung.STAND_NAME, _
                                   RefersTo:="=""" & modAnleitung.ANLEITUNG_STAND & """", _
                                   Visible:=False
            On Error GoTo 0

        Case 7
            If Not modSchutz.SchutzAktiv() Then modSchutz.Blattschutz_Umschalten

        Case 8
            If mSaveZahl > 0 Then
                modWochenplan.FastOn
                ws.Cells(CLng(mSaveZahl), "E").Value = mSaveText
                modWochenplan.FastOff
            End If

        Case 9, 10
            Set s = FormNachName(mSaveName)
            If Not s Is Nothing Then s.OnAction = mSaveText

    End Select
End Sub


'  Die n-te Schaltflaeche im Blatt "Steuerung", die ein Makro hat.
'
'  ACHTUNG - hier lag ein Fehler, den der Mutationstest am 05.09.2026
'  an sich selbst gefunden hat: die Sabotage 9 nimmt der ERSTEN Form
'  ihr Makro, und die Ruecknahme hat danach wieder "die erste Form MIT
'  Makro" gesucht. Das war nach der Sabotage aber eine ANDERE - die
'  Ruecknahme hat das gemerkte Makro also der zweiten Schaltflaeche
'  verpasst und die erste leer gelassen. Zwei kaputte Knoepfe statt
'  keinem.
'  Deshalb wird zum Zuruecknehmen ueber den NAMEN gegangen, nicht
'  ueber eine Eigenschaft, die die Sabotage selbst veraendert hat.
Private Function SteuerungsForm(ByVal welche As Long) As Shape
    Dim wsCtrl As Worksheet, i As Long, n As Long
    Set wsCtrl = Blatt(CTRL_SHEET)
    If wsCtrl Is Nothing Then Exit Function
    For i = 1 To wsCtrl.Shapes.Count
        If Len(wsCtrl.Shapes(i).OnAction) > 0 Then
            n = n + 1
            If n = welche Then
                Set SteuerungsForm = wsCtrl.Shapes(i)
                Exit Function
            End If
        End If
    Next i
End Function


Private Function FormNachName(ByVal nm As String) As Shape
    Dim wsCtrl As Worksheet
    Set wsCtrl = Blatt(CTRL_SHEET)
    If wsCtrl Is Nothing Or Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set FormNachName = wsCtrl.Shapes(nm)
    On Error GoTo 0
End Function


'  Eine Schaltflaeche ausloesen und pruefen, wo man danach landet.
Private Sub FokusProbe(ByVal bez As String, ByVal sollBlatt As String, _
                       ByVal welche As Long)
    Dim start As Worksheet, ist As String

    ' Bewusst woanders starten - sonst prueft man nichts.
    Set start = Blatt(modAnleitung.SHEET_HELP)
    If start Is Nothing Then Set start = Blatt(LB_SHEET)
    If start Is Nothing Then Exit Sub
    On Error Resume Next
    start.Activate
    On Error GoTo 0

    modWochenplan.ClearLastError
    Select Case welche
        Case 1: modKalender.UW_Und_Ferien_Generieren
        Case 2: modSteuerung.Plan_Fixieren
        Case 3: modSteuerung.Fixierung_Aufheben
        Case 4: If modSchutz.SchutzAktiv() Then modSchutz.Blattschutz_Umschalten
        Case 5: If Not modSchutz.SchutzAktiv() Then modSchutz.Blattschutz_Umschalten
        Case 6: modKopf.Kopf_Aktualisieren
        Case 7: modKalender.Ferien_Aus_Kalender_Vorschlagen
    End Select

    ist = ""
    On Error Resume Next
    ist = ActiveSheet.Name
    On Error GoTo 0

    Chk "Nach """ & bez & """ liegt der Fokus auf " & sollBlatt, _
        ist = sollBlatt, "tatsaechlich: " & ist
    If modWochenplan.LastError() <> "" Then
        Notiz bez & ": " & modWochenplan.LastError()
    End If
End Sub


'  Breite geteilt durch Hoehe des sichtbaren Logos.
Private Function LogoVerhaeltnis(ByVal ws As Worksheet) As Double
    Dim i As Long, s As Shape

    On Error Resume Next
    For i = 1 To ws.Shapes.Count
        Set s = ws.Shapes(i)
        If Left$(s.Name, Len(modKopf.LOGO_PRAEFIX)) = modKopf.LOGO_PRAEFIX Then
            If s.Visible Then
                If s.Height > 0 Then LogoVerhaeltnis = s.Width / s.Height
                Exit Function
            End If
        End If
    Next i
End Function


'=====================================================================
'  Hilfsmittel
'=====================================================================

'  Zustand des Plans als Text - Grundlage aller Rundlauf-Prüfungen.
Private Function Snap(ByVal ws As Worksheet) As String
    Dim r As Long, c As Long, lastRow As Long, s As String

    lastRow = modWochenplan.PlanLastRow(ws)
    For r = WP_FIRST_ROW To lastRow
        If modWochenplan.IsFerienRow(ws, r) Then
            s = s & r & "|FERIEN|" & T(ws.Cells(r, "B").Value) & vbLf
        Else
            s = s & r & "|-|" & T(ws.Cells(r, "B").Value) & "|" & T(ws.Cells(r, "E").Value)
            For c = 6 To 11
                s = s & "|" & T(ws.Cells(r, c).Value)
            Next c
            s = s & "|" & T(ws.Cells(r, "M").Value) & vbLf
        End If
    Next r
    Snap = s
End Function


'  Zellwert als Text, auch wenn eine Formel einen Fehler liefert
Private Function T(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        T = "#FEHLER"
    ElseIf IsEmpty(v) Then
        T = ""
    Else
        T = CStr(v)
    End If
End Function


Private Function HatGueltigkeit(ByVal c As Range) As Boolean
    Dim x As Long
    On Error Resume Next
    x = c.Validation.Type
    HatGueltigkeit = (Err.Number = 0)
    Err.Clear
End Function


'  Alle Texte einer Zeile zwischen zwei Spalten zu einem String.
'  Verbundene Zellen liefern ihren Wert in der linken oberen Zelle -
'  die uebrigen sind leer und fallen einfach weg.
Private Function ZeilenText(ByVal ws As Worksheet, ByVal r As Long, _
                            ByVal c1 As String, ByVal c2 As String) As String
    Dim c As Long, s As String, v As String

    On Error Resume Next
    For c = ws.Cells(r, c1).Column To ws.Cells(r, c2).Column
        v = Trim$(CStr(ws.Cells(r, c).Value))
        If Len(v) > 0 Then
            If Len(s) > 0 Then s = s & "   "
            s = s & v
        End If
    Next c
    ZeilenText = s
End Function


Private Function KopfWertTest(ByVal st As Worksheet, ByVal zeile As Long) As String
    On Error Resume Next
    If st Is Nothing Then Exit Function
    If zeile <= 0 Then Exit Function
    KopfWertTest = Trim$(CStr(st.Cells(zeile, modWochenplan.KopfCol() + 1).Value))
End Function


Private Function Blatt(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Blatt = ThisWorkbook.Worksheets(nm)
End Function


Private Sub Abschnitt(ByVal t As String)
    Zeile ""
    Zeile "----------------------------------------------------------"
    Zeile " " & t
    Zeile "----------------------------------------------------------"
End Sub


Private Sub Chk(ByVal bez As String, ByVal ok As Boolean, Optional ByVal detail As String = "")
    '  Fuer den Mutationstest: merken, ob GENAU der erwartete Check rot
    '  geworden ist. Ein Abschnitt, der irgendwo anders rot wird, zaehlt
    '  nicht - sonst wuerde eine Sabotage als "erkannt" durchgehen, die
    '  in Wahrheit einen ganz anderen Check umgeworfen hat.
    If Not ok And Len(mErwarte) > 0 Then
        If InStr(1, bez, mErwarte, vbTextCompare) > 0 Then mGetroffen = True
    End If

    If ok Then
        mOk = mOk + 1
        Zeile "  ok      " & bez
    Else
        mFail = mFail + 1
        Zeile "  FEHLER  " & bez
        If Len(detail) > 0 Then Zeile "          -> " & detail
        '  Auch fuer die Schlussmeldung merken: wenn der Bericht sich
        '  nicht schreiben laesst, ist das sonst die einzige Spur -
        '  und "1 durchgefallen" ohne Namen hilft niemandem.
        If Len(mFehlerListe) < 900 Then
            mFehlerListe = mFehlerListe & vbCrLf & "   - " & bez
            If Len(detail) > 0 Then mFehlerListe = mFehlerListe & vbCrLf & "     " & detail
        End If
    End If
End Sub


Private Sub Notiz(ByVal t As String)
    Dim teile As Variant, i As Long
    If Len(t) = 0 Then Exit Sub
    teile = Split(Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(teile) To UBound(teile)
        If Len(Trim$(CStr(teile(i)))) > 0 Then Zeile "  .       " & teile(i)
    Next i
End Sub


Private Sub Zeile(ByVal t As String)
    mRep = mRep & t & vbCrLf
End Sub


'  Der Bericht wird der Reihe nach an mehreren Stellen versucht.
'
'  Die alte Fassung kannte nur den Mappenordner und danach C:\Temp -
'  und beides ging beim Nutzer daneben: bei eingeschaltetem
'  AutoSpeichern meldet ThisWorkbook.Path die SharePoint-ADRESSE
'  (https://...), dorthin laesst sich mit Open ... For Output nicht
'  schreiben, und C:\Temp gibt es auf einem normalen Windows nicht.
'  Uebrig blieb "(Bericht konnte nicht geschrieben werden)" - und
'  damit kein Hinweis darauf, WELCHE Pruefung durchgefallen war.
Private Function BerichtSchreiben(Optional ByVal datei As String = "") As String
    Dim orte As Variant, i As Long, p As String

    If Len(datei) = 0 Then datei = BERICHT_DATEI

    orte = Array(ThisWorkbook.Path, Environ$("TEMP"), _
                 Environ$("USERPROFILE") & "\Documents", "C:\Temp")

    For i = LBound(orte) To UBound(orte)
        p = CStr(orte(i))
        If Len(p) > 0 And LCase$(Left$(p & "    ", 4)) <> "http" Then
            If Schreibversuch(p & "\" & datei) Then
                BerichtSchreiben = p & "\" & datei
                Exit Function
            End If
        End If
    Next i
    BerichtSchreiben = "(Bericht konnte nicht geschrieben werden)"
End Function


Private Function Schreibversuch(ByVal pfad As String) As Boolean
    Dim ff As Integer
    On Error Resume Next
    Err.Clear
    ff = FreeFile
    Open pfad For Output As #ff
    If Err.Number = 0 Then
        Print #ff, mRep
        Close #ff
    End If
    Schreibversuch = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
