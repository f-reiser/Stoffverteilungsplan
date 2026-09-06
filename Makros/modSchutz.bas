Attribute VB_Name = "modSchutz"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Blattschutz
'  ------------------------------------------------------------------
'  Ziel: Nur die Zellen, die wirklich von Hand gepflegt werden, sind
'  beschreibbar. Im Blatt "Wochenplan" laesst sich die Struktur ueber
'  die Excel-Oberflaeche gar nicht mehr aendern - Zeilen einfuegen,
'  loeschen, verschieben und sortieren koennen dort ausschliesslich die
'  Makros.
'
'  In "Lernbereiche" und "Einstellungen" duerfen Zeilen eingefuegt und
'  geloescht werden. Diese beiden Blaetter haben keine makrogepflegte
'  Struktur: die Lage aller Tabellen wird ueber die Beschriftungen
'  gesucht (modWochenplan.RefreshLayout).
'
'  Welche Zellen gesperrt werden, entscheidet UnlockRange automatisch:
'  ein Bereich wird freigegeben, Zellen MIT Formel darin bleiben
'  gesperrt. Dadurch sind auch die leeren Zeilen unterhalb der Daten
'  benutzbar, und niemand ueberschreibt versehentlich eine Formel.
'
'  Technik: Der Schutz wird fuer die Dauer einer Makroaktion
'  aufgehoben (Schutz_Aus / Schutz_An, aufgerufen von FastOn/FastOff).
'  WICHTIG fuer die Geschwindigkeit: Schutz_Aus nimmt optional das
'  EINE Blatt entgegen, das die Aktion anfasst. Dann wird nur dieses
'  Blatt entsperrt und die Mappenstruktur bleibt unberuehrt - das ist
'  um ein Vielfaches schneller als der volle Durchlauf ueber alle
'  Blaetter und macht das Verschieben von Zeilen wieder fluessig.
'  Den vollen Durchlauf brauchen nur Aktionen, die Blaetter anlegen
'  oder mehrere Blaetter gleichzeitig aendern.
'
'  Das Kennwort steht in modKonfig (SCHUTZ_PW) und wird von Updates
'  nicht angefasst.
'=====================================================================

' Reserve an leeren Zeilen, die unterhalb der Daten benutzbar bleiben
Private Const RESERVE_KALENDER As Long = 12
Private Const RESERVE_LB       As Long = 15
Private Const MIN_LB_ROW       As Long = 50

' Zustand des letzten Schutz_Aus
Private mKeepOpen As Collection      ' Blaetter, die schon vorher offen waren
Private mAnyProtected As Boolean     ' war ueberhaupt etwas geschuetzt?
Private mWbWasProtected As Boolean
Private mScopeName As String         ' "" = alle Blaetter, sonst nur dieses
Private mFehler As String            ' Blaetter, die sich nicht entsperren liessen


'=====================================================================
'  Schutz kurzzeitig aufheben / wiederherstellen (fuer die Makros)
'  onlySheet = Nothing  -> alle Blaetter und die Mappenstruktur
'  onlySheet = ein Blatt -> nur dieses Blatt (schnell)
'=====================================================================
Public Sub Schutz_Aus(Optional ByVal onlySheet As Worksheet = Nothing)
    Dim ws As Worksheet

    Set mKeepOpen = New Collection
    mAnyProtected = False
    mWbWasProtected = False
    mScopeName = ""
    mFehler = ""

    If Not onlySheet Is Nothing Then
        mScopeName = onlySheet.Name
        If onlySheet.ProtectContents Then
            mAnyProtected = True
            On Error Resume Next
            onlySheet.Unprotect SCHUTZ_PW
            On Error GoTo 0
            If onlySheet.ProtectContents Then
                ' Kennwort passt nicht - Blatt bleibt zu
                mAnyProtected = False
                MerkeFehler onlySheet.Name
            End If
        End If
        Exit Sub
    End If

    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect SCHUTZ_PW
        If Not ThisWorkbook.ProtectStructure Then mWbWasProtected = True
    End If
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then
            mAnyProtected = True
            On Error Resume Next
            ws.Unprotect SCHUTZ_PW
            On Error GoTo 0
            ' Blaetter mit fremdem Kennwort bleiben zu und werden
            ' behandelt, als waeren sie schon vorher offen gewesen.
            If ws.ProtectContents Then
                mKeepOpen.Add ws.Name
                MerkeFehler ws.Name
            End If
        Else
            mKeepOpen.Add ws.Name
        End If
    Next ws
End Sub


Public Sub Schutz_An()
    Dim ws As Worksheet

    If mAnyProtected Then
        If mScopeName <> "" Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(mScopeName)
            On Error GoTo 0
            If Not ws Is Nothing Then
                If Not ws.ProtectContents Then ProtectSheet ws
            End If
        Else
            ' Alles wieder schuetzen, was nicht ausdruecklich offen
            ' bleiben soll. Blaetter, die waehrend der Aktion neu
            ' entstanden sind, stehen nicht in mKeepOpen und werden
            ' dadurch mit geschuetzt.
            For Each ws In ThisWorkbook.Worksheets
                If Not InCollection(mKeepOpen, ws.Name) Then
                    If Not ws.ProtectContents Then ProtectSheet ws
                End If
            Next ws
        End If
    End If

    Set mKeepOpen = Nothing
    mAnyProtected = False
    mScopeName = ""

    If mWbWasProtected Then
        On Error Resume Next
        If Not ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Protect Password:=SCHUTZ_PW, Structure:=True, Windows:=False
        End If
        On Error GoTo 0
        mWbWasProtected = False
    End If
End Sub


Private Sub MerkeFehler(ByVal nm As String)
    If mFehler = "" Then mFehler = nm Else mFehler = mFehler & ", " & nm
End Sub


'  Namen der Blaetter, die sich beim letzten Schutz_Aus NICHT
'  entsperren liessen - fast immer ein falsches Kennwort.
Public Function EntsperrFehler() As String
    EntsperrFehler = mFehler
End Function


'  Prueft vor einer schreibenden Aktion, ob wirklich entsperrt ist,
'  und erklaert im Klartext, woran es sonst liegt.
Private Function SchreibbarPruefen() As Boolean
    If mFehler = "" Then
        SchreibbarPruefen = True
        Exit Function
    End If
    modWochenplan.FastOff
    modWochenplan.Info "Der Blattschutz liess sich nicht aufheben." & vbCrLf & vbCrLf & _
           "Betroffen: " & mFehler & vbCrLf & vbCrLf & _
           "Das Kennwort in der Konstante SCHUTZ_PW (Modul modKonfig) passt " & _
           "nicht zu dem Kennwort, mit dem diese Blätter geschützt wurden." & vbCrLf & vbCrLf & _
           "So kommst du weiter:" & vbCrLf & _
           "  1. Alt+F11, im Modul modKonfig das richtige Kennwort eintragen." & vbCrLf & _
           "  2. Zurück nach Excel, Einrichtung / Reparatur erneut ausführen." & vbCrLf & vbCrLf & _
           "Weisst du das alte Kennwort nicht mehr, hebe den Blattschutz einmal " & _
           "von Hand auf (Überprüfen, Blattschutz aufheben - auf jedem Blatt und " & _
           "über Überprüfen, Arbeitsmappe schützen), danach klappt die Einrichtung.", _
           vbExclamation, "Stoffverteilungsplan"
End Function


Private Function InCollection(ByVal col As Collection, ByVal nm As String) As Boolean
    Dim v As Variant
    If col Is Nothing Then Exit Function
    For Each v In col
        If CStr(v) = nm Then
            InCollection = True
            Exit Function
        End If
    Next v
End Function


'=====================================================================
'  Ein Blatt schuetzen
'  In "Lernbereiche" und "Einstellungen" bleiben Zeilen einfuegen und
'  loeschen erlaubt - dort gibt es keine makrogepflegte Struktur.
'=====================================================================
Private Sub ProtectSheet(ByVal ws As Worksheet)
    Dim rowsOk As Boolean
    rowsOk = (ws.Name = LB_SHEET Or ws.Name = SET_SHEET)

    On Error Resume Next
    ws.EnableSelection = xlNoRestrictions
    ws.Protect Password:=SCHUTZ_PW, _
               DrawingObjects:=False, Contents:=True, Scenarios:=False, _
               AllowFormattingCells:=False, _
               AllowFormattingColumns:=True, _
               AllowFormattingRows:=True, _
               AllowInsertingColumns:=False, _
               AllowInsertingRows:=rowsOk, _
               AllowInsertingHyperlinks:=False, _
               AllowDeletingColumns:=False, _
               AllowDeletingRows:=rowsOk, _
               AllowSorting:=False, _
               AllowFiltering:=False, _
               AllowUsingPivotTables:=False
    ws.EnableOutlining = True
    On Error GoTo 0
End Sub


'  DrawingObjects bleibt bewusst ungeschuetzt: nur so kann VBA die
'  Schaltflaechen weiter bewegen und notfalls neu anlegen. Loescht
'  jemand eine, baut EnsureButtons sie beim naechsten Klick in den
'  Wochenplan wieder auf.


'=====================================================================
'  Sperrkennzeichen setzen und alles schuetzen
'=====================================================================
Public Sub Blattschutz_Einrichten()
    Dim ws As Worksheet, wsCtrl As Worksheet

    If Not modWochenplan.LayoutReady() Then Exit Sub

    modWochenplan.FastOn
    If Not SchreibbarPruefen() Then Exit Sub
    On Error GoTo Fail

    ' Die Statuszeile sagt erst dann "aktiv", wenn wirklich geschuetzt
    ' ist - bricht etwas dazwischen ab, steht dort weiterhin "AUS".
    modWochenplan.SetStep "Statuszeile zuruecksetzen"
    StatusSchreiben False

    modWochenplan.SetStep "Sperrkennzeichen Wochenplan"
    LockWochenplan modWochenplan.WpSheet()

    modWochenplan.SetStep "Sperrkennzeichen Einstellungen"
    LockEinstellungen modWochenplan.SetSheet()

    modWochenplan.SetStep "Sperrkennzeichen Lernbereiche"
    LockLernbereiche SheetOrNothing(LB_SHEET)

    modWochenplan.SetStep "Sperrkennzeichen Infoblaetter"
    LockAll SheetOrNothing(CTRL_SHEET)
    ' Die Auswahlliste fuer das Papierformat muss bedienbar bleiben -
    ' sie ist die einzige Eingabe auf diesem Blatt.
    FreigebenNachName modKopf.PDF_FORMAT_NAME
    LockAll SheetOrNothing(modAnleitung.SHEET_HELP)

    ' Zuerst alle Blaetter ausser der Steuerung schuetzen, dann die
    ' Statuszeile schreiben, dann die Steuerung selbst.
    modWochenplan.SetStep "Blaetter schuetzen"
    Set wsCtrl = SheetOrNothing(CTRL_SHEET)
    For Each ws In ThisWorkbook.Worksheets
        If wsCtrl Is Nothing Then
            ProtectSheet ws
        ElseIf Not ws Is wsCtrl Then
            ProtectSheet ws
        End If
    Next ws

    modWochenplan.SetStep "Statuszeile"
    StatusSchreiben True
    If Not wsCtrl Is Nothing Then ProtectSheet wsCtrl

    modWochenplan.SetStep "Mappenstruktur schuetzen"
    On Error Resume Next
    ThisWorkbook.Protect Password:=SCHUTZ_PW, Structure:=True, Windows:=False
    On Error GoTo Fail

    modWochenplan.FastOff
    Exit Sub
Fail:
    modWochenplan.ReportError "Einrichten des Blattschutzes"
End Sub


Private Sub Blattschutz_Aufheben()
    Dim ws As Worksheet

    If modWochenplan.Frage("Blattschutz aufheben? Danach können Zeilen im " & _
              WP_SHEET & " auch von Hand eingefügt und gelöscht werden - das " & _
              "bringt Unterrichtswochen, Ferienzeilen und Formeln durcheinander.", _
              vbExclamation + vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

    On Error GoTo Fail

    On Error Resume Next
    ThisWorkbook.Unprotect SCHUTZ_PW
    On Error GoTo Fail

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect SCHUTZ_PW
        On Error GoTo Fail
    Next ws

    Set mKeepOpen = Nothing
    mAnyProtected = False
    mWbWasProtected = False
    mScopeName = ""
    StatusSchreiben False

    ' Zurueck auf die Steuerung - dort sitzt die Schaltflaeche. Das
    ' muss ausdruecklich passieren: dieser Weg laeuft ohne FastOn und
    ' FastOff, also greift deren Blattmerker hier nicht.
    modWochenplan.GotoSheet CTRL_SHEET
    modWochenplan.Info "Der Blattschutz ist aufgehoben."
    Exit Sub
Fail:
    modWochenplan.ReportError "Aufheben des Blattschutzes"
End Sub


Public Sub Blattschutz_Umschalten()
    If SchutzAktiv() Then
        Blattschutz_Aufheben
    Else
        Blattschutz_Einrichten
        modWochenplan.GotoSheet CTRL_SHEET
        If SchutzAktiv() Then modWochenplan.Info "Der Blattschutz ist aktiv."
    End If
End Sub


'  Wird beim Öffnen der Mappe aufgerufen (Workbook_Open). Der Schutz
'  selbst wird mit der Datei gespeichert; nur die Bedienbarkeit der
'  Gliederung (die Plus/Minus-Schaltflaeche ueber den Spalten P bis S)
'  muss nach jedem Oeffnen neu erlaubt werden.
Public Sub Blattschutz_Nachziehen()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then ws.EnableOutlining = True
    Next ws
End Sub


Public Function SchutzAktiv() As Boolean
    Dim ws As Worksheet
    Set ws = modWochenplan.WpSheet()
    If ws Is Nothing Then Exit Function
    SchutzAktiv = ws.ProtectContents
End Function


'=====================================================================
'  Sperrkennzeichen je Blatt
'=====================================================================

'  Hebt die Sperre fuer einen Bereich auf - Zellen MIT Formel bleiben
'  aber gesperrt. Dadurch muss nirgends aufgezaehlt werden, welche
'  Spalten Formeln enthalten, und leere Zeilen unterhalb der Daten
'  bleiben benutzbar.
Private Sub UnlockRange(ByVal rng As Range)
    Dim c As Range
    If rng Is Nothing Then Exit Sub
    rng.Locked = False
    For Each c In rng.Cells
        If c.HasFormula Then c.Locked = True
    Next c
End Sub


Private Sub LockWochenplan(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long, a As Long
    Dim ferien As Boolean
    If ws Is Nothing Then Exit Sub

    ws.Cells.Locked = True
    lastRow = modWochenplan.PlanLastRow(ws)

    ' E (UW), F bis K (Inhalt, Erledigt, Stunde) und M (Notizen)
    ' freigeben - aber blockweise zwischen den Ferienzeilen.
    '
    ' WICHTIG: In einer Ferienzeile sind B bis M zu EINER Zelle
    ' verbunden. Ein Bereich wie E2:K40 wuerde diese Verbindung nur
    ' teilweise ueberdecken, und genau das quittiert Excel mit
    ' "Die Locked-Eigenschaft des Range-Objektes kann nicht festgelegt
    ' werden" (Laufzeitfehler 1004). Deshalb wird je zusammenhaengendem
    ' Block von Planzeilen einzeln freigegeben. Die Ferienzeilen selbst
    ' bleiben gesperrt - das sind sie durch ws.Cells.Locked = True oben
    ' bereits.
    a = 0
    For r = WP_FIRST_ROW To lastRow + 1
        If r > lastRow Then
            ferien = True
        Else
            ferien = modWochenplan.IsFerienRow(ws, r)
        End If

        If ferien Then
            If a > 0 Then
                UnlockRange ws.Range("E" & a & ":K" & (r - 1))
                UnlockRange ws.Range("M" & a & ":M" & (r - 1))
                a = 0
            End If
        Else
            If a = 0 Then a = r
        End If
    Next r
End Sub


Private Sub LockEinstellungen(ByVal st As Worksheet)
    Dim h As Long, kc As Long
    If st Is Nothing Then Exit Sub

    st.Cells.Locked = True

    ' Kopfblock: die Werte in Spalte B, ohne Formeln und ohne den
    ' Fixierungs-Schalter
    h = modWochenplan.WeekHeadRow()
    If h > 1 Then UnlockRange st.Range("B1:B" & (h - 1))
    st.Cells(modWochenplan.FixRow(), "B").Locked = True

    ' Kalendertabelle inklusive einiger leerer Zeilen darunter, damit
    ' sich weitere Schulwochen ergaenzen lassen. Die Rechenspalten G
    ' und H sind in den gefuellten Zeilen Formeln und bleiben gesperrt,
    ' in den leeren Zeilen sind sie frei.
    UnlockRange st.Range(st.Cells(modWochenplan.WeekFirstRow(), "A"), _
                         st.Cells(modWochenplan.WeekLastRow() + RESERVE_KALENDER, "H"))

    ' Ferientabelle (hat von Haus aus freie Zeilen)
    UnlockRange st.Range(st.Cells(modWochenplan.FerFirstRow(), FER_COL_VON), _
                         st.Cells(modWochenplan.FerLastRow(), FER_COL_NAME))

    ' Die vier Kopfangaben: jeweils die Zelle RECHTS NEBEN der
    ' Beschriftung. Frueher stand hier kr+1 bis kr+4 - das galt fuer die
    ' alte Fassung, in der ueber den vier Werten noch eine Ueberschrift
    ' stand. Seit die Beschriftungen selbst gesucht werden, faengt der
    ' Block bei kr an, und kr+1..kr+4 haette ausgerechnet "Fach"
    ' gesperrt gelassen. Jede Zeile wird deshalb einzeln freigegeben -
    ' die vier duerfen auch auseinanderstehen.
    kc = modWochenplan.KopfCol()
    If kc > 0 Then
        KopfZelleFrei st, modWochenplan.KopfZeileFach(), kc
        KopfZelleFrei st, modWochenplan.KopfZeileKlasse(), kc
        KopfZelleFrei st, modWochenplan.KopfZeileSchule(), kc
        KopfZelleFrei st, modWochenplan.KopfZeileLehrer(), kc
    End If

    ' Schulliste: Name und Logo-Datei bleiben aenderbar, damit sich
    ' eine weitere Schule ergaenzen laesst. Ein paar Leerzeilen mit
    ' dazu, sonst kaeme man an die naechste freie Zeile nicht heran.
    If modWochenplan.SchulRow() > 0 Then
        UnlockRange st.Range(st.Cells(modWochenplan.SchulRow() + 1, modWochenplan.SchulCol()), _
                             st.Cells(modWochenplan.SchulLastRow() + 5, modWochenplan.SchulCol() + 1))
    End If

    ' Die Stundenliste bleibt bewusst gesperrt: an ihr haengen die fuenf
    ' Regeln der bedingten Formatierung und die Gueltigkeitsliste der
    ' Spalte K. Wer dort Werte aendert, bekommt Farben und Auswahl
    ' auseinander - und die Makros schreiben die Liste ohnehin nach.
End Sub


Private Sub KopfZelleFrei(ByVal st As Worksheet, ByVal zeile As Long, _
                          ByVal kc As Long)
    If zeile <= 0 Then Exit Sub
    UnlockRange st.Cells(zeile, kc + 1)
End Sub


Private Sub LockLernbereiche(ByVal ws As Worksheet)
    Dim lastRow As Long
    If ws Is Nothing Then Exit Sub

    ws.Cells.Locked = True
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row + RESERVE_LB
    If lastRow < MIN_LB_ROW Then lastRow = MIN_LB_ROW

    ' Ganzer Datenbereich einschliesslich der leeren Zeilen darunter.
    ' Die Rechenspalten F bis I und die Summenzeile sind Formeln und
    ' bleiben dadurch automatisch gesperrt; in leeren Zeilen sind sie
    ' frei, damit sich die Formeln von oben herunterziehen lassen.
    ' KEINE feste 2: ueber der Ueberschriftenzeile steht der Titelblock,
    ' und der darf nicht bearbeitbar werden.
    If lastRow < modWochenplan.LB_FIRST_ROW() Then Exit Sub
    UnlockRange ws.Range("A" & modWochenplan.LB_FIRST_ROW() & ":J" & lastRow)
End Sub


'  Eine ueber einen benannten Bereich bezeichnete Zelle entsperren.
Private Sub FreigebenNachName(ByVal nm As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Locked = False
    On Error GoTo 0
End Sub


Private Sub LockAll(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Locked = True
End Sub


Private Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(nm)
End Function


Private Sub StatusSchreiben(ByVal aktiv As Boolean)
    Dim ws As Worksheet
    Set ws = SheetOrNothing(CTRL_SHEET)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If aktiv Then
        ws.Range(CTRL_STATUS_SCHUTZ).Value = "Blattschutz: aktiv"
        ws.Range(CTRL_STATUS_SCHUTZ).Font.Color = modWochenplan.FARBE_LEISE
    Else
        ws.Range(CTRL_STATUS_SCHUTZ).Value = "Blattschutz: AUS - bitte wieder einschalten"
        ws.Range(CTRL_STATUS_SCHUTZ).Font.Color = RGB(168, 96, 96)
    End If
End Sub
