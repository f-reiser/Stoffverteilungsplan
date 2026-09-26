Attribute VB_Name = "modWochenplan"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - Kernmodul fuer das Tabellenblatt "Wochenplan"
'  ------------------------------------------------------------------
'  Vier Schaltflaechen links neben der Markierung:
'      Pfeil hoch   Block nach oben schieben (nur Inhalt F:K und M,
'                   im fixierten Zustand zusaetzlich Spalte B)
'      +            so viele Zeilen einfuegen wie markiert sind
'      -            markierte Zeilen loeschen
'      Pfeil runter Block nach unten schieben
'
'  Ferien-Zeilen  - Zeilen mit dem Kennzeichen "F" in Spalte N.
'    Bei ihnen sind B:M zu einer Zelle verbunden. Sie werden beim
'    Verschieben uebersprungen, beim Loeschen ausgelassen und bei
'    Formeln/Gueltigkeit ausgespart. Erzeugt werden sie von
'    modKalender.UW_Und_Ferien_Generieren.
'
'  Fixierung - steht der Schalter im Blatt "Einstellungen" auf WAHR,
'    sind die Referenz-Codes in Spalte B feste Werte. Sie wandern beim
'    Verschieben mit, bleiben beim Loeschen unveraendert und werden
'    fuer neue Zeilen fortlaufend neu vergeben (hoechste vorhandene
'    Nummer + 1).
'
'  WICHTIG 1: Dieses Modul fasst die bedingte Formatierung NICHT an -
'  weder lesend noch schreibend. Sieben der acht Regeln sind
'  x14-Erweiterungsregeln; ein Zugriff auf die FormatConditions-
'  Auflistung bringt Excel zum Absturz. Stattdessen werden neue Zeilen
'  ueber "kopierte Zellen einfuegen" erzeugt - dabei dehnt Excel die
'  Regelbereiche selbst mit aus, auch am Tabellenende. Beim Loeschen
'  zieht Excel die Bereiche von sich aus nach.
'
'  WICHTIG 2: Im Blatt "Einstellungen" ist KEINE Zeile fest verdrahtet.
'  Die Lage der Kalendertabelle, der Stundenliste, des Fixierungs-
'  Schalters und der Ferientabelle wird zur Laufzeit anhand der
'  Beschriftungen gesucht (siehe RefreshLayout). Wer dort Zeilen
'  einfuegt oder Bloecke verschiebt, muss die Makros nicht anfassen.
'=====================================================================

' ---------------- Konfiguration --------------------------------------
Public Const WP_SHEET      As String = "Wochenplan"
Public Const LB_SHEET      As String = "Lernbereiche"
Public Const SET_SHEET     As String = "Einstellungen"

' Beschriftungen, an denen die Kopfzeilen der Datenblaetter erkannt
' werden. Darueber liegt seit 09/2026 der Titelblock (Fach, Klasse,
' Schule, Lehrkraft, Logo) - die erste Datenzeile ist deshalb NICHT
' mehr fest Zeile 2.
Public Const LBL_WPHEAD    As String = "Referenz Code"
Public Const LBL_LBHEAD    As String = "Referenzcode"

' Markierung der Ferien-/Sonderzeilen im Wochenplan
Public Const MARK_COL      As String = "N"

'  Kennzeichen einer Sonderzeile in Spalte N. Bewusst nur "F": die
'  Tabelle heisst inzwischen "Ferien / eingeplante Fehlzeiten" und
'  nimmt auch Fortbildungen oder Praktika auf - "FERIEN" stand da also
'  falsch. Geschrieben wird es in der Hintergrundfarbe der Zelle, ist
'  also nicht zu sehen; gebraucht wird es nur intern.
Public Const MARK_TAG      As String = "F"

'  Was frueher dort stand. Aeltere Mappen sollen weiterlaufen, ohne
'  dass ihre Ferienzeilen ploetzlich als normale Planzeilen gelten -
'  das wuerde beim Formelneuaufbau in die verbundenen Zellen
'  schreiben und mit Laufzeitfehler 1004 enden.
Public Const MARK_TAG_ALT  As String = "FERIEN"

' Ferientabelle im Blatt "Einstellungen" (Spalten; die Zeilen ergeben
' sich aus der Kopfzeile der Kalendertabelle)
Public Const FER_COL_VON   As String = "I"
Public Const FER_COL_BIS   As String = "J"
Public Const FER_COL_NAME  As String = "K"
Public Const FER_ROWS      As Long = 20

' Beschriftungen, an denen sich das Blatt "Einstellungen" erkennen laesst
Public Const LBL_YEAR      As String = "Schuljahresstart"
Public Const LBL_STDWOCHE  As String = "Stunden pro Woche"
Public Const LBL_WEEKSUM   As String = "verfügbare Wochen insgesamt"
Public Const LBL_KLWOCHEN  As String = "verfügbare Klassen-Wochen"
Public Const LBL_FIX       As String = "Stoffverteilungsplan fixiert"
Public Const LBL_WEEKHEAD  As String = "Schulwoche"
Public Const LBL_STDLIST   As String = "Auswahlbereich"
Public Const LBL_FERVON    As String = "Ferien von"

'  Der Fixierungs-Schalter steht als Text da, nicht als WAHR/FALSCH -
'  "ja"/"nein" liest sich fuer den Anwender einfach naheliegender.
'  IsPlanFixed nimmt beides an, damit aeltere Mappen weiterlaufen.
Public Const FIX_JA   As String = "ja"
Public Const FIX_NEIN As String = "nein"
'  Reine Beschriftung der Spalte K - wird NICHT zur Erkennung benutzt.
'  Bewusst generisch: dort duerfen auch Fortbildungen, Praktika und
'  andere eingeplante Fehlzeiten stehen.
Public Const LBL_FERNAME   As String = "Ferien / eingeplante Fehlzeiten"
Public Const LBL_FACH      As String = "Fach"
Public Const LBL_KLASSE    As String = "Klasse"
Public Const LBL_SCHULE    As String = "Schule"
Public Const LBL_LEHRER    As String = "Lehrkraft"
Public Const LBL_SCHULLISTE As String = "Schulliste"
'  Beschriftung der Summenzeile im Blatt "Lernbereiche" (Spalte D)
Public Const LBL_LBSUMME   As String = "Summe"

' Suchbereich im Blatt "Einstellungen" fuer die Layout-Erkennung
Private Const SCAN_ROWS As Long = 200
Private Const SCAN_COLS As Long = 30

Private Const MOVE_COL_A1 As String = "F"     ' Verschiebeblock 1 von
Private Const MOVE_COL_A2 As String = "K"     ' Verschiebeblock 1 bis
Private Const MOVE_COL_B1 As String = "M"     ' Verschiebeblock 2 (Notizen)

'  Wie weit der FILTER in "Lernbereiche" ueber die erste Datenzeile
'  hinaus schaut. Bisher endete der Bereich fest bei Zeile 50.
Private Const LB_FILTER_RESERVE As Long = 48

Private Const CF_COL_FIRST As String = "B"
Private Const CF_COL_LAST  As String = "M"

Private Const BTN_UP   As String = "wpBtnUp"
Private Const BTN_ADD  As String = "wpBtnAdd"
Private Const BTN_DEL  As String = "wpBtnDel"
Private Const BTN_DOWN As String = "wpBtnDown"

Private Const DIAG_FALLBACK_DIR As String = "C:\Temp"

' Blatt, auf das die laufende Aktion beschraenkt ist (Nothing = alle)

' ---------------- Laufzeitzustand ------------------------------------
Private mCalcSave As XlCalculation
Private mEventSave As Boolean
Private mScreenSave As Boolean
Private mAlertSave As Boolean
Private mBusy As Boolean
Private mStep As String
Private mProbeGelaufen As Boolean
Private mScopeSheet As Worksheet

' ---------------- Erkanntes Layout des Blattes "Einstellungen" -------
Private mLayoutOk As Boolean
Private mRowYear As Long          ' Zeile "Schuljahresstart"
Private mRowWeekSum As Long       ' Zeile "verfuegbare Wochen insgesamt"
Private mRowStdWoche As Long      ' Zeile "Stunden pro Woche"
Private mRowKlWochen As Long      ' Zeile "verfuegbare Klassen-Wochen"
Private mRowFix As Long           ' Zeile des Fixierungs-Schalters
Private mWeekHead As Long         ' Kopfzeile der Kalendertabelle
Private mWeekFirst As Long        ' erste Datenzeile der Kalendertabelle
Private mWeekLast As Long         ' letzte Datenzeile der Kalendertabelle
Private mStdCol As Long           ' Spalte der Stunden-Bezeichnungen
Private mStdFirst As Long
Private mStdLast As Long
Private mKopfFach As Long         ' Zeile der Beschriftung "Fach"
Private mKopfCol As Long          ' Spalte der Kopf-Beschriftungen
Private mKopfKlasse As Long
Private mKopfSchule As Long
Private mKopfLehrer As Long
Private mSchulRow As Long         ' Zeile der Beschriftung "Schulliste"
Private mSchulCol As Long
Private mSchulLast As Long        ' letzte Zeile der Schulliste

  ' Kopfzeilen der Datenblaetter (0 = noch nicht ermittelt)
Private mWpHead As Long
Private mLbHead As Long

' Aktives Blatt vor der Makroaktion - wird von FastOff wiederhergestellt
Private mActiveSave As Object

' ---------------- Stiller Modus (Selbsttest) -------------------------
' Im stillen Modus erscheinen keine Dialoge: Rueckfragen werden mit OK
' beantwortet, Meldungen nur gesammelt. So kann modSelbsttest die
' Abläufe durchspielen, ohne dass jemand klicken muss.
Private mQuiet As Boolean
Private mLastInfo As String
Private mLastError As String


'=====================================================================
'  Layout des Blattes "Einstellungen" ermitteln
'=====================================================================
Public Function RefreshLayout(Optional ByVal force As Boolean = True) As Boolean
    Dim st As Worksheet, r As Long, c As Long, v As String
    Dim a As Variant

    If mLayoutOk And Not force Then
        RefreshLayout = True
        Exit Function
    End If

    mLayoutOk = False
    Set st = SetSheet()
    If st Is Nothing Then Exit Function

    ' Der ganze Suchbereich wird in EINEM Zugriff als Array geholt.
    ' Zellweises Lesen waere hier rund tausendmal langsamer - und
    ' RefreshLayout laeuft bei jedem Formelneuaufbau.
    On Error Resume Next
    a = st.Range(st.Cells(1, 1), st.Cells(SCAN_ROWS, SCAN_COLS)).Value
    On Error GoTo 0
    If Not IsArray(a) Then Exit Function

    ' --- Beschriftungen im Kopfblock (Spalte A) ----------------------
    mRowYear = 0: mRowWeekSum = 0: mRowFix = 0: mWeekHead = 0
    mRowStdWoche = 0: mRowKlWochen = 0
    For r = 1 To SCAN_ROWS
        v = Txt(a, r, 1)
        If v <> "" Then
            If v = LCase$(LBL_YEAR) Then mRowYear = r
            If v = LCase$(LBL_STDWOCHE) Then mRowStdWoche = r
            If v = LCase$(LBL_WEEKSUM) Then mRowWeekSum = r
            If v = LCase$(LBL_KLWOCHEN) Then mRowKlWochen = r
            If v = LCase$(LBL_FIX) Then mRowFix = r
            If v = LCase$(LBL_WEEKHEAD) And mWeekHead = 0 Then mWeekHead = r
        End If
    Next r

    If mRowYear = 0 Then mRowYear = 1
    If mRowStdWoche = 0 Then mRowStdWoche = 2
    If mRowWeekSum = 0 Then mRowWeekSum = 3
    If mRowKlWochen = 0 Then mRowKlWochen = 4
    If mWeekHead = 0 Then Exit Function            ' ohne Kalender geht nichts

    ' --- Datenzeilen der Kalendertabelle -----------------------------
    mWeekFirst = mWeekHead + 1
    mWeekLast = mWeekFirst - 1
    For r = mWeekFirst To SCAN_ROWS
        If IsNum(a, r, 1) Then
            mWeekLast = r
        Else
            Exit For
        End If
    Next r
    If mWeekLast < mWeekFirst Then Exit Function

    ' --- Fixierungs-Schalter -----------------------------------------
    If mRowFix = 0 Then mRowFix = FreeLabelRow(st)

    ' --- Stundenliste ("Auswahlbereich fuer Stunde") -----------------
    mStdCol = 0: mStdFirst = 0: mStdLast = 0
    FindLabel a, LBL_STDLIST, r, c
    If r > 0 Then
        mStdCol = c + 1
        mStdFirst = r + 1
        mStdLast = r
        Do While r < SCAN_ROWS
            r = r + 1
            If IsNum(a, r, c) Then
                mStdLast = r
            Else
                Exit Do
            End If
        Loop
    End If
    If mStdCol = 0 Or mStdLast < mStdFirst Then
        mStdCol = 13: mStdFirst = 2: mStdLast = 6     ' Rueckfallebene M2:M6
    End If

    ' --- Angaben fuer den Titelblock ---------------------------------
    ' Die vier Beschriftungen werden EINZELN gesucht. Der Nutzer darf
    ' sie also verschieben, solange Fach/Klasse/Schule/Lehrkraft in
    ' derselben Spalte untereinander stehen und der Wert rechts daneben.
    mKopfFach = 0: mKopfKlasse = 0: mKopfSchule = 0: mKopfLehrer = 0
    mKopfCol = 0
    FindLabel a, LBL_FACH, mKopfFach, c
    If mKopfFach > 0 Then mKopfCol = c
    FindLabel a, LBL_KLASSE, mKopfKlasse, c
    FindLabel a, LBL_SCHULE, mKopfSchule, c
    FindLabel a, LBL_LEHRER, mKopfLehrer, c

    ' --- Schulliste ---------------------------------------------------
    mSchulRow = 0: mSchulCol = 0: mSchulLast = 0
    FindLabel a, LBL_SCHULLISTE, mSchulRow, mSchulCol
    If mSchulRow > 0 Then
        For r = mSchulRow + 1 To SCAN_ROWS
            If Txt(a, r, mSchulCol) <> "" Then
                mSchulLast = r
            Else
                Exit For
            End If
        Next r
    End If

    mLayoutOk = True
    RefreshLayout = True
End Function


'  Zellinhalt aus dem Array als kleingeschriebener Text (fehlertolerant)
Private Function Txt(ByRef a As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error Resume Next
    If IsError(a(r, c)) Then Exit Function
    If IsEmpty(a(r, c)) Then Exit Function
    Txt = LCase$(Trim$(CStr(a(r, c))))
End Function


Private Function IsNum(ByRef a As Variant, ByVal r As Long, ByVal c As Long) As Boolean
    On Error Resume Next
    If IsError(a(r, c)) Then Exit Function
    If IsEmpty(a(r, c)) Then Exit Function
    IsNum = IsNumeric(a(r, c))
End Function


'  Sucht eine Beschriftung im Array und liefert Zeile und Spalte
'  (0/0, wenn sie nicht vorkommt). Verglichen wird der Zellanfang,
'  damit ein angehaengter Zusatztext nicht stoert.
Private Sub FindLabel(ByRef a As Variant, ByVal lbl As String, _
                      ByRef foundRow As Long, ByRef foundCol As Long)
    Dim r As Long, c As Long, v As String, n As Long, low As String
    foundRow = 0: foundCol = 0
    n = Len(lbl)
    low = LCase$(lbl)
    For c = 1 To SCAN_COLS
        For r = 1 To SCAN_ROWS
            v = Txt(a, r, c)
            If Len(v) >= n Then
                If Left$(v, n) = low Then
                    foundRow = r
                    foundCol = c
                    Exit Sub
                End If
            End If
        Next r
    Next c
End Sub


Private Function FreeLabelRow(ByVal st As Worksheet) As Long
    Dim r As Long
    For r = 4 To mWeekHead - 1
        If Len(Trim$(CStr(st.Cells(r, 1).Value))) = 0 Then
            FreeLabelRow = r
            Exit Function
        End If
    Next r
    FreeLabelRow = 5
End Function


Private Sub NeedLayout()
    If Not mLayoutOk Then RefreshLayout True
End Sub


' ---- oeffentliche Auskuenfte ueber das erkannte Layout --------------
Public Function WeekHeadRow() As Long
    NeedLayout
    WeekHeadRow = mWeekHead
End Function

Public Function WeekFirstRow() As Long
    NeedLayout
    WeekFirstRow = mWeekFirst
End Function

Public Function WeekLastRow() As Long
    NeedLayout
    WeekLastRow = mWeekLast
End Function

Public Function YearRow() As Long
    NeedLayout
    YearRow = mRowYear
End Function

Private Function StdWocheRow() As Long
    NeedLayout
    StdWocheRow = mRowStdWoche
End Function

Private Function KlassenWochenRow() As Long
    NeedLayout
    KlassenWochenRow = mRowKlWochen
End Function

Private Function WeekSumRow() As Long
    NeedLayout
    WeekSumRow = mRowWeekSum
End Function

Public Function FixRow() As Long
    NeedLayout
    FixRow = mRowFix
End Function

Public Function FixCellAddr() As String
    FixCellAddr = "B" & FixRow()
End Function

Public Function FerHeadRow() As Long
    FerHeadRow = WeekHeadRow()
End Function

Public Function FerFirstRow() As Long
    FerFirstRow = WeekHeadRow() + 1
End Function

Public Function FerLastRow() As Long
    FerLastRow = WeekHeadRow() + FER_ROWS
End Function

'=====================================================================
'  Kopfzeilen der Datenblaetter
'  Ueber den Daten steht seit 09/2026 ein Titelblock. Die erste
'  Datenzeile ergibt sich deshalb aus der Zeile mit den
'  Spaltenueberschriften und ist nicht mehr fest verdrahtet.
'  WP_FIRST_ROW ist bewusst eine FUNKTION mit dem alten Namen - so
'  bleiben alle Aufrufstellen unveraendert gueltig.
'=====================================================================
Public Function WP_HEAD_ROW() As Long
    If mWpHead = 0 Then mWpHead = SucheKopfzeile(WpSheet(), "B", LBL_WPHEAD)
    WP_HEAD_ROW = mWpHead
End Function

Public Function WP_FIRST_ROW() As Long
    WP_FIRST_ROW = WP_HEAD_ROW() + 1
End Function

Public Function LB_HEAD_ROW() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LB_SHEET)
    On Error GoTo 0
    If mLbHead = 0 Then mLbHead = SucheKopfzeile(ws, "A", LBL_LBHEAD)
    LB_HEAD_ROW = mLbHead
End Function

Public Function LB_FIRST_ROW() As Long
    LB_FIRST_ROW = LB_HEAD_ROW() + 1
End Function

'  Muss nach dem Einfuegen oder Loeschen von Kopfzeilen aufgerufen
'  werden, damit die naechste Abfrage neu sucht.
'=====================================================================
'  F A R B E N
'  ------------------------------------------------------------------
'  Alle Blautoene der Mappe liegen auf dem Farbton des Schullogos
'  (#1468AB, 207 Grad). Vorher hatte der Tabellenkopf zwar denselben
'  Farbton, war aber deutlich blasser - neben dem kraeftigen Logo
'  wirkte das muffig. Jetzt ist er gleich gesaettigt und nur klar
'  dunkler, damit das Logo der hellste Punkt bleibt.
'
'  Bewusst FUNKTIONEN und keine Konstanten: eine Const kann RGB()
'  nicht aufrufen, und eine von Hand ausgerechnete Farbzahl war schon
'  zweimal falsch, ohne dass man es gesehen haette.
'=====================================================================
Public Function FARBE_KOPF() As Long          ' #16507E  Tabellenkopf
    FARBE_KOPF = RGB(22, 80, 126)
End Function
Public Function FARBE_KOPF_HELL() As Long     ' #5F8EB4  zweiter Kopfton
    FARBE_KOPF_HELL = RGB(95, 142, 180)
End Function
Public Function FARBE_BALKEN() As Long        ' #4E697E  Abschnittsbalken
    FARBE_BALKEN = RGB(78, 105, 126)
End Function
Public Function FARBE_HELL() As Long          ' #E1E9EF  helle Flaeche
    FARBE_HELL = RGB(225, 233, 239)
End Function
Public Function FARBE_TEXT() As Long          ' #48505D  Fliesstext
    FARBE_TEXT = RGB(72, 80, 93)
End Function
Public Function FARBE_LEISE() As Long         ' #6E7A91  Nebentext
    FARBE_LEISE = RGB(110, 122, 145)
End Function


Public Sub ResetKopfzeilen()
    mWpHead = 0
    mLbHead = 0
End Sub

Private Function SucheKopfzeile(ByVal ws As Worksheet, ByVal spalte As String, _
                                ByVal lbl As String) As Long
    Dim r As Long, v As String
    SucheKopfzeile = 1                      ' Rueckfallebene: alter Aufbau
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    For r = 1 To 30
        v = ""
        v = LCase$(Trim$(CStr(ws.Cells(r, spalte).Value)))
        If v = LCase$(lbl) Then
            SucheKopfzeile = r
            Exit Function
        End If
    Next r
End Function


'=====================================================================
'  Angaben fuer den Titelblock (Blatt "Einstellungen")
'=====================================================================
Private Function KopfRow() As Long
    NeedLayout
    KopfRow = mKopfFach
End Function

Public Function KopfCol() As Long
    NeedLayout
    KopfCol = mKopfCol
End Function

Public Function KopfZeileFach() As Long
    NeedLayout
    KopfZeileFach = mKopfFach
End Function

Public Function KopfZeileKlasse() As Long
    NeedLayout
    KopfZeileKlasse = mKopfKlasse
End Function

Public Function KopfZeileSchule() As Long
    NeedLayout
    KopfZeileSchule = mKopfSchule
End Function

Public Function KopfZeileLehrer() As Long
    NeedLayout
    KopfZeileLehrer = mKopfLehrer
End Function

Public Function SchulRow() As Long
    NeedLayout
    SchulRow = mSchulRow
End Function

Public Function SchulCol() As Long
    NeedLayout
    SchulCol = mSchulCol
End Function

Public Function SchulLastRow() As Long
    NeedLayout
    SchulLastRow = mSchulLast
End Function

Public Function StdCol() As Long
    NeedLayout
    StdCol = mStdCol
End Function

Public Function StdFirstRow() As Long
    NeedLayout
    StdFirstRow = mStdFirst
End Function

Private Function StdLastRow() As Long
    NeedLayout
    StdLastRow = mStdLast
End Function

'  Public, seit modPruefung und der Selbsttest die Adresse der
'  Stundenliste brauchen: an ihr haengen die Gueltigkeitsliste der
'  Spalte K und die Farbregeln der bedingten Formatierung.
Public Function StdListAddress() As String
    Dim st As Worksheet
    NeedLayout
    Set st = SetSheet()
    If st Is Nothing Then Exit Function
    StdListAddress = "Einstellungen!" & _
        st.Range(st.Cells(mStdFirst, mStdCol), st.Cells(mStdLast, mStdCol)).Address(True, True)
End Function

Public Function LayoutInfo() As String
    NeedLayout
    LayoutInfo = "Kalender  : Kopf " & mWeekHead & ", Daten " & mWeekFirst & "-" & mWeekLast & vbCrLf & _
                 "Ferien    : Kopf " & FerHeadRow() & ", Daten " & FerFirstRow() & "-" & FerLastRow() & _
                 " (Spalten " & FER_COL_VON & "/" & FER_COL_BIS & "/" & FER_COL_NAME & ")" & vbCrLf & _
                 "Fixierung : " & FixCellAddr() & vbCrLf & _
                 "Stunden   : " & StdListAddress() & vbCrLf & _
                 "Jahr / Summe: B" & mRowYear & " / B" & mRowWeekSum
End Function


'=====================================================================
'  Formeltexte (US-Schreibweise, wie sie Range.Formula erwartet)
'  Jede Funktion bekommt die Zeilennummer, fuer die die Formel gilt -
'  dadurch koennen die Formeln auch fuer Bloecke geschrieben werden,
'  die nicht bei Zeile 2 beginnen (Ferienzeilen unterbrechen sie).
'  Die Bezuege auf "Einstellungen" stammen aus RefreshLayout.
'=====================================================================
Private Function ColRef(ByVal col As String) As String
    ' z. B. "Einstellungen!$B$9:$B$46"
    ColRef = "Einstellungen!$" & col & "$" & mWeekFirst & ":$" & col & "$" & mWeekLast
End Function

Private Function CellRef(ByVal col As String, ByVal r As Long) As String
    CellRef = "Einstellungen!$" & col & "$" & r
End Function

Private Function FrmA(ByVal r As Long) As String
    FrmA = "=IF($E" & r & "="""","""",IF(COUNTIF(P" & r & ":S" & r & ",""!""),""!"",""""))"
End Function
Private Function FrmB(ByVal r As Long) As String
    FrmB = "=IF($F" & r & "="""","""",$F" & r & "&""-""&TEXT(COUNTA($F$" & _
           WP_FIRST_ROW & ":$F" & r & "),""00""))"
End Function
Private Function FrmV(ByVal r As Long) As String
    FrmV = "=IFERROR(MATCH($E" & r & "," & ColRef("A") & ",0),0)"
End Function
Private Function FrmC(ByVal r As Long) As String
    FrmC = "=IF($V" & r & "=0,"""",TEXT(INDEX(" & ColRef("B") & ",$V" & r & "),""TT.MM"")" & _
           "&"" - ""&TEXT(INDEX(" & ColRef("C") & ",$V" & r & "),""TT.MM""))"
End Function
Private Function FrmD(ByVal r As Long) As String
    FrmD = "=IF($V" & r & "=0,"""",_xlfn.ISOWEEKNUM(DATE(IF(MONTH(INDEX(" & ColRef("B") & ",$V" & r & "))>=9," & _
           CellRef("B", mRowYear) & "," & CellRef("B", mRowYear) & "+1)," & _
           "MONTH(INDEX(" & ColRef("B") & ",$V" & r & "))," & _
           "DAY(INDEX(" & ColRef("B") & ",$V" & r & ")))))"
End Function
Private Function FrmL(ByVal r As Long) As String
    FrmL = "=IF($V" & r & "=0,"""",IF(AND(INDEX(" & ColRef("E") & ",$V" & r & "),INDEX(" & ColRef("F") & ",$V" & r & "))," & _
           """1/2 Woche""&CHAR(10)&""1/2 Klasse""," & _
           "IF(INDEX(" & ColRef("E") & ",$V" & r & "),""1/2 Woche""," & _
           "IF(INDEX(" & ColRef("F") & ",$V" & r & "),""1/2 Klasse"",""""))))"
End Function
Private Function FrmU(ByVal r As Long) As String
    Dim z1 As Long, z2 As Long, a As String, b As String, c As String

    ' KEINE feste Zeile 2: ueber der Ueberschriftenzeile der
    ' Lernbereiche steht der Titelblock, der die Daten nach unten
    ' schiebt. Der Bereich beginnt deshalb bei LB_FIRST_ROW und reicht
    ' bewusst grosszuegig darueber hinaus, damit beliebig viele
    ' Lernbereiche moeglich sind. Leerzeilen und die Summenzeile werden
    ' vom Filter nicht getroffen, weil dort "von"/"bis" leer sind.
    z1 = LB_FIRST_ROW()
    z2 = z1 + LB_FILTER_RESERVE
    a = "Lernbereiche!$A$" & z1 & ":$A$" & z2
    b = "Lernbereiche!$B$" & z1 & ":$B$" & z2
    c = "Lernbereiche!$C$" & z1 & ":$C$" & z2

    FrmU = "=IFERROR(INDEX(_xlfn._xlws.FILTER(" & a & "," & _
           "(" & b & "<=$E" & r & ")*(" & c & ">=$E" & r & "),""""),1),"""")"
End Function
Private Function FrmP(ByVal r As Long) As String
    FrmP = "=IF(OR($V" & r & "=0,$E" & r & ">" & CellRef("B", mRowWeekSum) & "),""""," & _
           "IF(INDEX(" & ColRef("D") & ",$V" & r & "),"""",""!""))"
End Function
Private Function FrmQ(ByVal r As Long) As String
    FrmQ = "=IF($E" & r & "="""","""",IF(OR($U" & r & "=$F" & r & ",_xlfn.REGEXTEST($F" & r & ",""P$"")),"""",""!""))"
End Function
Private Function FrmR(ByVal r As Long) As String
    FrmR = "=IF($E" & r & "="""","""",IF($E" & r & "<=" & CellRef("B", mRowWeekSum) & ","""",""!""))"
End Function
Private Function FrmS(ByVal r As Long, ByVal rPrev As Long) As String
    ' rPrev = letzte Planzeile oberhalb, die KEINE Ferienzeile ist
    FrmS = "=IF($E" & r & "="""","""",IF($E" & r & ">$E" & rPrev & ","""",""!""))"
End Function


'=====================================================================
'  Oeffentliche Einstiegspunkte
'=====================================================================
Private Sub Setup_Wochenplan()
    Dim ws As Worksheet
    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub
    If Not LayoutReady() Then Exit Sub

    FastOn ws
    On Error GoTo Fail
    SetStep "RebuildAll"
    RebuildAll ws, PlanLastRow(ws)
    SetStep "EnsureButtons"
    EnsureButtons ws
    SetStep "HideButtons"
    HideButtons ws
    FastOff
    Info "Der Wochenplan wurde eingerichtet." & vbCrLf & vbCrLf & _
           "Markiere eine oder mehrere Zeilen - links am Rand erscheinen " & _
           "die vier Schaltflächen.", vbInformation, "Stoffverteilungsplan"
    Exit Sub
Fail:
    ReportError "Aufbau"
End Sub


Public Sub Wochenplan_Aktualisieren()
    Dim ws As Worksheet
    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub
    If Not LayoutReady() Then Exit Sub
    FastOn ws
    On Error GoTo Fail
    RebuildAll ws, PlanLastRow(ws)
    FastOff
    Exit Sub
Fail:
    ReportError "Aktualisieren"
End Sub


'  Prueft das Layout und meldet sich, wenn die Kalendertabelle im Blatt
'  "Einstellungen" nicht gefunden wird.
Public Function LayoutReady() As Boolean
    If RefreshLayout(True) Then
        LayoutReady = True
        Exit Function
    End If
    Info "Im Blatt '" & SET_SHEET & "' wurde die Kalendertabelle nicht gefunden." & vbCrLf & vbCrLf & _
           "In Spalte A muss eine Zelle mit der Beschriftung '" & LBL_WEEKHEAD & "' stehen; " & _
           "darunter folgen die durchnummerierten Schulwochen.", _
           vbExclamation, "Stoffverteilungsplan"
End Function


'  Legt die vier Schaltflaechen am Zeilenrand an, falls sie fehlen,
'  und blendet sie aus. Wird beim Einrichten aufgerufen, damit sie
'  garantiert vorhanden sind, bevor der Blattschutz greift.
Public Sub EnsureRowButtons()
    Dim ws As Worksheet
    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    EnsureButtons ws
    HideButtons ws
End Sub


Public Sub UpdateRowButtons(Optional ByVal Target As Range)
    Dim ws As Worksheet, r1 As Long, r2 As Long, lastRow As Long

    If mBusy Then Exit Sub
    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then
        If TypeName(Selection) <> "Range" Then Exit Sub
        Set Target = Selection
    End If
    On Error GoTo CleanFail

    lastRow = PlanLastRow(ws)
    If Not GetBlock(ws, Target, lastRow, r1, r2) Then
        HideButtons ws
        Exit Sub
    End If

    ' Auf einer reinen Ferienzeile koennen die Schaltflaechen nichts
    ' ausrichten - dann werden sie gar nicht erst gezeigt.
    If Not HatInhaltszeile(ws, r1, r2) Then
        HideButtons ws
        Exit Sub
    End If

    EnsureButtons ws
    PlaceButtons ws, r1, r2
    Exit Sub
CleanFail:
End Sub


'---------------------------------------------------------------------
'  Wird aus dem Blattmodul bei jeder Zelleingabe aufgerufen.
'  Im fixierten Zustand bekommt eine Zeile, in der ein Lehrplan-Code
'  (Spalte F) neu eingetragen wird und die noch keinen Referenz-Code
'  besitzt, sofort einen neuen eindeutigen Code.
'---------------------------------------------------------------------
Public Sub OnPlanChange(ByVal Target As Range)
    Dim ws As Worksheet, isect As Range, c As Range, lastRow As Long
    Dim nxt As Long, didChange As Boolean

    If mBusy Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Not IsPlanFixed() Then Exit Sub

    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub
    If Not Target.Parent Is ws Then Exit Sub

    On Error GoTo CleanFail
    lastRow = PlanLastRow(ws)
    Set isect = Application.Intersect(Target, ws.Range("F" & WP_FIRST_ROW & ":F" & lastRow))
    If isect Is Nothing Then Exit Sub

    nxt = NextRefNumber(ws, lastRow)

    ' FastOn schaltet Ereignisse ab und hebt den Blattschutz auf -
    ' Spalte B ist im fixierten Zustand fuer die Eingabe gesperrt.
    FastOn ws
    For Each c In isect.Cells
        If Not IsFerienRow(ws, c.Row) Then
            If Len(Trim$(CStr(c.Value))) > 0 Then
                If Len(Trim$(CStr(ws.Cells(c.Row, "B").Value))) = 0 Then
                    ws.Cells(c.Row, "B").Value = Trim$(CStr(c.Value)) & "-" & Format$(nxt, "00")
                    nxt = nxt + 1
                    didChange = True
                End If
            End If
        End If
    Next c
    FastOff
    Exit Sub
CleanFail:
    FastOff
End Sub


'=====================================================================
'  Klick-Handler
'=====================================================================
Public Sub wpBtnUp_Click()
    Dim r1 As Long, r2 As Long
    If Not ReadBlock(r1, r2) Then Exit Sub
    MoveBlock r1, r2, -1
End Sub

Public Sub wpBtnDown_Click()
    Dim r1 As Long, r2 As Long
    If Not ReadBlock(r1, r2) Then Exit Sub
    MoveBlock r1, r2, 1
End Sub

Public Sub wpBtnAdd_Click()
    Dim r1 As Long, r2 As Long
    If Not ReadBlock(r1, r2) Then Exit Sub
    InsertRowsBelow r1, r2
End Sub

Public Sub wpBtnDel_Click()
    Dim r1 As Long, r2 As Long
    If Not ReadBlock(r1, r2) Then Exit Sub
    DeleteRowsBlock r1, r2
End Sub


'=====================================================================
'  Aktion 1 - Zeilenblock verschieben
'  Verschoben wird nur der Inhalt (F:K, M und im fixierten Zustand B).
'  Gerechnet wird auf der Liste der Inhaltszeilen, Ferienzeilen werden
'  dadurch uebersprungen und bleiben an ihrem Platz.
'=====================================================================
Public Sub MoveBlock(ByVal r1 As Long, ByVal r2 As Long, ByVal d As Long)
    Dim ws As Worksheet, lastRow As Long
    Dim cr() As Long, nRows As Long
    Dim i1 As Long, i2 As Long, cnt As Long, k As Long
    Dim vals() As Variant
    Dim fixMode As Boolean

    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = PlanLastRow(ws)
    nRows = ContentRows(ws, lastRow, cr)
    If nRows = 0 Then Exit Sub
    If Not BlockIndex(cr, nRows, r1, r2, i1, i2) Then Beep: Exit Sub

    If d < 0 Then
        If i1 <= 1 Then Beep: Exit Sub
    Else
        If i2 >= nRows Then Beep: Exit Sub
    End If

    FastOn ws
    On Error GoTo Fail

    cnt = i2 - i1 + 1
    fixMode = IsPlanFixed()
    ReDim vals(1 To cnt + 1, 1 To 9)

    SetStep "Werte lesen"
    For k = 1 To cnt
        ReadRowVals ws, cr(i1 + k - 1), vals, k
    Next k
    If d < 0 Then
        ReadRowVals ws, cr(i1 - 1), vals, cnt + 1
    Else
        ReadRowVals ws, cr(i2 + 1), vals, cnt + 1
    End If

    SetStep "Werte schreiben"
    If d < 0 Then
        ' Reihenfolge in den Zeilen cr(i1-1) .. cr(i2):  Block, dann Nachbar
        For k = 1 To cnt
            WriteRowVals ws, cr(i1 - 2 + k), vals, k, fixMode
        Next k
        WriteRowVals ws, cr(i2), vals, cnt + 1, fixMode
    Else
        ' Reihenfolge in den Zeilen cr(i1) .. cr(i2+1):  Nachbar, dann Block
        WriteRowVals ws, cr(i1), vals, cnt + 1, fixMode
        For k = 1 To cnt
            WriteRowVals ws, cr(i1 + k), vals, k, fixMode
        Next k
    End If

    ' Ein Verschieben aendert nur Werte, nicht die Struktur: alle
    ' Formeln sind zeilenrelativ und stimmen weiterhin. Der komplette
    ' Neuaufbau waere hier reine Wartezeit - nur die Referenz-Codes
    ' muessen im fixierten Zustand nachgezogen werden.
    SetStep "Referenz-Codes"
    AssignMissingCodes ws, lastRow

    SetStep "Auswahl setzen"
    SelectBlock ws, cr(i1 + d), cr(i2 + d)
    FastOff
    UpdateRowButtons ws.Range(ws.Cells(cr(i1 + d), CF_COL_FIRST), ws.Cells(cr(i2 + d), CF_COL_LAST))
    Exit Sub
Fail:
    ReportError "Verschieben"
End Sub


Private Sub ReadRowVals(ByVal ws As Worksheet, ByVal r As Long, _
                        ByRef vals() As Variant, ByVal idx As Long)
    Dim c As Long
    vals(idx, 1) = ws.Cells(r, CF_COL_FIRST).Value          ' B
    For c = 6 To 11                                          ' F..K
        vals(idx, c - 4) = ws.Cells(r, c).Value
    Next c
    vals(idx, 8) = ws.Cells(r, MOVE_COL_B1).Value            ' M
    ' Die Zeilenhoehe gehoert zum Inhalt: eine Zeile mit viel Text
    ' braucht ihre Hoehe auch an der neuen Stelle.
    vals(idx, 9) = ws.Rows(r).RowHeight
End Sub


Private Sub WriteRowVals(ByVal ws As Worksheet, ByVal r As Long, _
                         ByRef vals() As Variant, ByVal idx As Long, _
                         ByVal withRef As Boolean)
    Dim c As Long
    If withRef Then ws.Cells(r, CF_COL_FIRST).Value = vals(idx, 1)
    For c = 6 To 11
        ws.Cells(r, c).Value = vals(idx, c - 4)
    Next c
    ws.Cells(r, MOVE_COL_B1).Value = vals(idx, 8)
    If IsNumeric(vals(idx, 9)) Then ws.Rows(r).RowHeight = vals(idx, 9)
End Sub


'=====================================================================
'  Aktion 2 - Neue Zeilen unterhalb des Blocks einfuegen
'=====================================================================
Public Sub InsertRowsBelow(ByVal r1 As Long, ByVal r2 As Long)
    Dim ws As Worksheet, lastRow As Long, n As Long, i As Long
    Dim cr() As Long, nRows As Long, i1 As Long, i2 As Long, anchor As Long

    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = PlanLastRow(ws)
    nRows = ContentRows(ws, lastRow, cr)
    If nRows = 0 Then Exit Sub
    If Not BlockIndex(cr, nRows, r1, r2, i1, i2) Then
        Info "Bitte mindestens eine normale Planzeile markieren " & _
               "(Ferienzeilen zählen nicht).", vbExclamation, "Stoffverteilungsplan"
        Exit Sub
    End If

    n = i2 - i1 + 1
    anchor = cr(i2)

    If n > 10 Then
        If Frage(n & " neue Zeilen einfügen?", vbQuestion + vbOKCancel, _
                  "Stoffverteilungsplan") <> vbOK Then Exit Sub
    End If

    FastOn ws
    On Error GoTo Fail

    ' Zeile "anchor" kopieren und als "kopierte Zellen" einfuegen: dabei
    ' uebernimmt Excel selbst Formate, Kontrollkaestchen, Zeilenhoehe,
    ' Gueltigkeitsliste UND bedingte Formatierung auf die neuen Zeilen
    ' und dehnt die Regelbereiche mit aus - auch am Tabellenende.
    SetStep "Zeilen einfuegen"
    Application.CutCopyMode = False
    For i = 1 To n
        ws.Rows(anchor).Copy
        ws.Rows(anchor + 1).Insert Shift:=xlDown
    Next i
    Application.CutCopyMode = False

    SetStep "Neue Zeilen leeren"
    ws.Range(ws.Cells(anchor + 1, "A"), ws.Cells(anchor + n, "V")).ClearContents
    ws.Rows(anchor + 1).Resize(n).RowHeight = ws.Rows(anchor).RowHeight
    For i = anchor + 1 To anchor + n
        ws.Cells(i, "J").Value = False
    Next i

    RebuildAll ws, lastRow + n

    SetStep "Auswahl setzen"
    SelectBlock ws, anchor + 1, anchor + n
    FastOff
    UpdateRowButtons ws.Range(ws.Cells(anchor + 1, CF_COL_FIRST), ws.Cells(anchor + n, CF_COL_LAST))
    Exit Sub
Fail:
    ReportError "Einfuegen"
End Sub


'=====================================================================
'  Aktion 3 - Markierte Zeilen loeschen (Ferienzeilen bleiben stehen)
'=====================================================================
Public Sub DeleteRowsBlock(ByVal r1 As Long, ByVal r2 As Long)
    Dim ws As Worksheet, lastRow As Long, n As Long, msg As String
    Dim cr() As Long, nRows As Long, i1 As Long, i2 As Long
    Dim k As Long, sel As Long

    Set ws = WpSheet()
    If ws Is Nothing Then Exit Sub

    lastRow = PlanLastRow(ws)
    nRows = ContentRows(ws, lastRow, cr)
    If nRows = 0 Then Exit Sub
    If Not BlockIndex(cr, nRows, r1, r2, i1, i2) Then
        Info "In der Markierung ist keine normale Planzeile enthalten." & vbCrLf & _
               "Ferienzeilen werden über den Kalender verwaltet und hier nicht gelöscht.", _
               vbExclamation, "Stoffverteilungsplan"
        Exit Sub
    End If

    n = i2 - i1 + 1
    If nRows - n < 1 Then
        Info "Es muss mindestens eine Planzeile übrig bleiben.", vbExclamation, _
               "Stoffverteilungsplan"
        Exit Sub
    End If

    If n = 1 Then
        msg = "Zeile " & cr(i1) & " endgültig löschen?"
    Else
        msg = n & " Planzeilen (Zeile " & cr(i1) & " bis " & cr(i2) & ") endgültig löschen?"
    End If
    If Frage(msg & vbCrLf & vbCrLf & "Das lässt sich nicht rückgängig machen.", _
              vbExclamation + vbOKCancel + vbDefaultButton2, _
              "Stoffverteilungsplan") <> vbOK Then Exit Sub

    FastOn ws
    On Error GoTo Fail

    SetStep "Zeilen loeschen"
    ' von unten nach oben, damit die noch offenen Zeilennummern gueltig bleiben
    For k = i2 To i1 Step -1
        ws.Rows(cr(k)).Delete Shift:=xlUp
    Next k

    sel = cr(i1)
    lastRow = PlanLastRow(ws)
    If sel > lastRow Then sel = lastRow
    If sel < WP_FIRST_ROW Then sel = WP_FIRST_ROW

    RebuildAll ws, lastRow

    SetStep "Auswahl setzen"
    SelectBlock ws, sel, sel
    FastOff
    UpdateRowButtons ws.Range(ws.Cells(sel, CF_COL_FIRST), ws.Cells(sel, CF_COL_LAST))
    Exit Sub
Fail:
    ReportError "Loeschen"
End Sub


'=====================================================================
'  Neuaufbau bzw. Nachziehen
'=====================================================================
Public Sub RebuildAll(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < WP_FIRST_ROW Then lastRow = WP_FIRST_ROW
    If Not RefreshLayout(True) Then Exit Sub
    MarkierungenAuffrischen ws, lastRow
    RebuildFormulas ws, lastRow
    RebuildValidation ws, lastRow
    AssignMissingCodes ws, lastRow
    ' Die bedingte Formatierung bleibt bewusst unangetastet - siehe
    ' Hinweis im Modulkopf.
End Sub


'  Alle vorhandenen Sonderzeilen auf das aktuelle Kennzeichen bringen.
'  Damit reicht "Einrichtung / Reparatur", um eine aeltere Mappe
'  umzustellen - man muss die Ferienzeilen nicht neu erzeugen.
Private Sub MarkierungenAuffrischen(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = WP_FIRST_ROW To lastRow
        If IsFerienRow(ws, r) Then MarkierungSetzen ws, r
    Next r
End Sub


Private Sub RebuildFormulas(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim segA() As Long, segB() As Long, nSeg As Long, s As Long
    Dim a As Long, b As Long, prevContent As Long
    Dim c As Variant
    Dim fixMode As Boolean

    fixMode = IsPlanFixed()
    nSeg = BuildSegments(ws, lastRow, segA, segB)

    SetStep "Formel O1"
    ws.Range("O" & WP_HEAD_ROW()).Formula = _
        "=IF(COUNTIF(P" & WP_FIRST_ROW & ":S" & lastRow & ",""!""),""Warnungen"","""")"

    If nSeg = 0 Then Exit Sub

    prevContent = 0
    For s = 1 To nSeg
        a = segA(s)
        b = segB(s)

        For Each c In Array("A", "C", "D", "L", "P", "Q", "R", "S", "U", "V")
            SetStep "Formeln leeren, Spalte " & c & " (Zeile " & a & "-" & b & ")"
            ws.Range(CStr(c) & a & ":" & CStr(c) & b).ClearContents
        Next c
        If Not fixMode Then
            SetStep "Formeln leeren, Spalte B (Zeile " & a & "-" & b & ")"
            ws.Range("B" & a & ":B" & b).ClearContents
        End If

        SetStep "Formel Spalte A": ws.Range("A" & a & ":A" & b).Formula = FrmA(a)
        If Not fixMode Then
            SetStep "Formel Spalte B": ws.Range("B" & a & ":B" & b).Formula = FrmB(a)
        End If
        SetStep "Formel Spalte V": ws.Range("V" & a & ":V" & b).Formula = FrmV(a)
        SetStep "Formel Spalte C": ws.Range("C" & a & ":C" & b).Formula = FrmC(a)
        SetStep "Formel Spalte D": ws.Range("D" & a & ":D" & b).Formula = FrmD(a)
        SetStep "Formel Spalte L": ws.Range("L" & a & ":L" & b).Formula = FrmL(a)
        SetStep "Formel Spalte U": SetFormulaSmart ws.Range("U" & a & ":U" & b), FrmU(a)
        SetStep "Formel Spalte P": ws.Range("P" & a & ":P" & b).Formula = FrmP(a)
        SetStep "Formel Spalte Q": ws.Range("Q" & a & ":Q" & b).Formula = FrmQ(a)
        SetStep "Formel Spalte R": ws.Range("R" & a & ":R" & b).Formula = FrmR(a)

        SetStep "Formel Spalte S (Zeile " & a & "-" & b & ")"
        If prevContent > 0 Then
            ws.Range("S" & a).Formula = FrmS(a, prevContent)
        End If
        If b > a Then
            ws.Range("S" & (a + 1) & ":S" & b).Formula = FrmS(a + 1, a)
        End If

        prevContent = b
    Next s
End Sub


Private Sub SetFormulaSmart(ByVal rng As Range, ByVal frm As String)
    On Error Resume Next
    Err.Clear
    rng.Formula2 = frm
    If Err.Number <> 0 Then
        Err.Clear
        rng.Formula = frm
    End If
    On Error GoTo 0
End Sub


Private Sub RebuildValidation(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim segA() As Long, segB() As Long, nSeg As Long, s As Long
    Dim lst As String

    nSeg = BuildSegments(ws, lastRow, segA, segB)
    If nSeg = 0 Then Exit Sub

    lst = StdListAddress()
    If Len(lst) = 0 Then Exit Sub

    SetStep "Gueltigkeitsliste Spalte K"
    On Error Resume Next
    For s = 1 To nSeg
        With ws.Range("K" & segA(s) & ":K" & segB(s)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=" & lst
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = True
            .ErrorTitle = "Ungültige Eingabe"
            .ErrorMessage = "Bitte einen Wert aus der Liste wählen."
        End With
    Next s
    On Error GoTo 0
End Sub


'=====================================================================
'  Referenz-Codes im fixierten Zustand
'=====================================================================
Public Function IsPlanFixed() As Boolean
    Dim st As Worksheet, v As Variant
    Set st = SetSheet()
    If st Is Nothing Then Exit Function
    NeedLayout
    On Error Resume Next
    v = st.Range(FixCellAddr()).Value
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsPlanFixed = CBool(v)
    Else
        IsPlanFixed = (UCase$(Trim$(CStr(v))) = "WAHR" Or UCase$(Trim$(CStr(v))) = "TRUE" _
                       Or Trim$(CStr(v)) = "1" Or UCase$(Trim$(CStr(v))) = "JA")
    End If
End Function


'  Hoechste bisher vergebene laufende Nummer in Spalte B + 1
Public Function NextRefNumber(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim r As Long, t As String, p As Long, n As Long, mx As Long
    mx = 0
    For r = WP_FIRST_ROW To lastRow
        If Not IsFerienRow(ws, r) Then
            t = Trim$(CStr(ws.Cells(r, CF_COL_FIRST).Value))
            p = InStrRev(t, "-")
            If p > 0 Then
                t = Mid$(t, p + 1)
                If Len(t) > 0 Then
                    If IsNumeric(t) Then
                        n = CLng(Val(t))
                        If n > mx Then mx = n
                    End If
                End If
            End If
        End If
    Next r
    NextRefNumber = mx + 1
End Function


Private Sub AssignMissingCodes(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long, nxt As Long, f As String
    If Not IsPlanFixed() Then Exit Sub
    SetStep "Referenz-Codes ergaenzen"
    nxt = NextRefNumber(ws, lastRow)
    For r = WP_FIRST_ROW To lastRow
        If Not IsFerienRow(ws, r) Then
            f = Trim$(CStr(ws.Cells(r, "F").Value))
            If Len(f) > 0 Then
                If Len(Trim$(CStr(ws.Cells(r, CF_COL_FIRST).Value))) = 0 Then
                    ws.Cells(r, CF_COL_FIRST).Value = f & "-" & Format$(nxt, "00")
                    nxt = nxt + 1
                End If
            End If
        End If
    Next r
End Sub


'=====================================================================
'  Ferienzeilen / Inhaltszeilen
'=====================================================================
Public Function IsFerienRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim v As String
    On Error Resume Next
    v = UCase$(Trim$(CStr(ws.Cells(r, MARK_COL).Value)))
    IsFerienRow = (v = MARK_TAG Or v = MARK_TAG_ALT)
End Function


'---------------------------------------------------------------------
'  Das Kennzeichen setzen und unsichtbar machen: Schriftfarbe gleich
'  Hintergrundfarbe der Zelle. Bringt gleichzeitig aeltere Mappen von
'  "FERIEN" auf "F".
'---------------------------------------------------------------------
Public Sub MarkierungSetzen(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    With ws.Cells(r, MARK_COL)
        .Value = MARK_TAG
        If .Interior.Pattern = xlNone Then
            .Font.Color = RGB(255, 255, 255)
        Else
            .Font.Color = .Interior.Color
        End If
    End With
    On Error GoTo 0
End Sub


'  Liefert die Zeilennummern aller Planzeilen, die keine Ferienzeilen
'  sind, in cr(1..n). Rueckgabewert = n.
Public Function ContentRows(ByVal ws As Worksheet, ByVal lastRow As Long, _
                            ByRef cr() As Long) As Long
    Dim r As Long, n As Long, cap As Long
    cap = lastRow - WP_FIRST_ROW + 1
    If cap < 1 Then cap = 1
    ReDim cr(1 To cap)
    n = 0
    For r = WP_FIRST_ROW To lastRow
        If Not IsFerienRow(ws, r) Then
            n = n + 1
            cr(n) = r
        End If
    Next r
    ContentRows = n
End Function


'  Bildet einen markierten Zeilenbereich auf Indizes der Inhaltsliste ab
Private Function BlockIndex(ByRef cr() As Long, ByVal nRows As Long, _
                            ByVal r1 As Long, ByVal r2 As Long, _
                            ByRef i1 As Long, ByRef i2 As Long) As Boolean
    Dim k As Long
    i1 = 0: i2 = 0
    For k = 1 To nRows
        If cr(k) >= r1 And cr(k) <= r2 Then
            If i1 = 0 Then i1 = k
            i2 = k
        End If
    Next k
    BlockIndex = (i1 > 0)
End Function


'  Zerlegt WP_FIRST_ROW..lastRow in zusammenhaengende Bloecke ohne
'  Ferienzeilen. Rueckgabewert = Anzahl der Bloecke.
Private Function BuildSegments(ByVal ws As Worksheet, ByVal lastRow As Long, _
                               ByRef segA() As Long, ByRef segB() As Long) As Long
    Dim r As Long, n As Long, inSeg As Boolean, cap As Long
    cap = lastRow - WP_FIRST_ROW + 1
    If cap < 1 Then cap = 1
    ReDim segA(1 To cap)
    ReDim segB(1 To cap)
    n = 0
    inSeg = False
    For r = WP_FIRST_ROW To lastRow
        If IsFerienRow(ws, r) Then
            inSeg = False
        Else
            If Not inSeg Then
                n = n + 1
                segA(n) = r
                inSeg = True
            End If
            segB(n) = r
        End If
    Next r
    BuildSegments = n
End Function


'=====================================================================
'  Schaltflaechen am Zeilenrand
'=====================================================================

' Formtypen werden ueber die benannten Office-Konstanten geholt - so
' kann es keine Verwechslung durch abweichende Nummerierung geben
' (msoShapeMathPlus statt einer geratenen Zahl).
Private Function ShTriangle() As Long
    ShTriangle = msoShapeIsoscelesTriangle
End Function

Private Function ShPlus() As Long
    ShPlus = msoShapeMathPlus
End Function

Private Function ShMinus() As Long
    ShMinus = msoShapeMathMinus
End Function

Private Function ShCross() As Long
    ShCross = msoShapeCross
End Function

Private Function ShRect() As Long
    ShRect = msoShapeRectangle
End Function


Private Sub EnsureButtons(ByVal ws As Worksheet)
    MakeButton ws, BTN_UP, ShTriangle(), ShTriangle(), 0, "wpBtnUp_Click", _
               FARBE_HELL, FARBE_LEISE
    MakeButton ws, BTN_ADD, ShPlus(), ShCross(), 0, "wpBtnAdd_Click", _
               RGB(219, 237, 222), RGB(93, 143, 102)
    MakeButton ws, BTN_DEL, ShMinus(), ShRect(), 0, "wpBtnDel_Click", _
               RGB(247, 223, 223), RGB(168, 96, 96)
    MakeButton ws, BTN_DOWN, ShTriangle(), ShTriangle(), 180, "wpBtnDown_Click", _
               FARBE_HELL, FARBE_LEISE
End Sub


Private Function MakeButton(ByVal ws As Worksheet, ByVal nm As String, _
                            ByVal shType As Long, ByVal shFallback As Long, _
                            ByVal rot As Single, ByVal macro As String, _
                            ByVal fillCol As Long, ByVal lineCol As Long) As Shape
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    On Error GoTo 0

    If s Is Nothing Then
        On Error Resume Next
        Set s = ws.Shapes.AddShape(shType, 100, 100, 12, 12)
        If s Is Nothing Then
            Err.Clear
            Set s = ws.Shapes.AddShape(shFallback, 100, 100, 12, 12)
        End If
        Err.Clear
        On Error GoTo 0
        If s Is Nothing Then Exit Function

        s.Name = nm
        With s
            .Placement = xlFreeFloating
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = fillCol
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = lineCol
            .Line.Weight = 0.75
            .Shadow.Visible = msoFalse
            .Rotation = rot
        End With
    End If

    s.OnAction = macro
    Set MakeButton = s
End Function


Private Sub PlaceButtons(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Dim topP As Double, botP As Double, colL As Double, colW As Double
    Dim bw As Double, bh As Double, gap As Double, totH As Double, y As Double, x As Double
    Dim tag As String
    Dim hoch As Boolean, runter As Boolean, n As Long, i As Long

    ' Nach oben bzw. nach unten laesst sich nur verschieben, wenn es
    ' dort ueberhaupt noch eine Planzeile gibt. Am oberen und unteren
    ' Ende wird die jeweilige Schaltflaeche deshalb gar nicht erst
    ' gezeigt - vorher lag sie da und tat nichts.
    hoch = (InhaltszeileVor(ws, r1) > 0)
    runter = (InhaltszeileNach(ws, r2) > 0)
    n = 2
    If hoch Then n = n + 1
    If runter Then n = n + 1

    topP = ws.Rows(r1).Top
    botP = ws.Rows(r2).Top + ws.Rows(r2).Height
    colL = ws.Columns("A").Left
    colW = ws.Columns("A").Width

    gap = 1#
    bw = colW - 4
    If bw > 13 Then bw = 13
    If bw < 3 Then bw = 3

    bh = ((botP - topP) - 3 - (n - 1) * gap) / n
    If bh > 13 Then bh = 13
    If bh < 4 Then bh = 4

    totH = n * bh + (n - 1) * gap
    y = topP + ((botP - topP) - totH) / 2
    x = colL + (colW - bw) / 2

    tag = r1 & ":" & r2

    i = 0
    If hoch Then
        PlaceOne ws, BTN_UP, x, y + i * (bh + gap), bw, bh, tag
        i = i + 1
    Else
        HideOne ws, BTN_UP, tag
    End If

    PlaceOne ws, BTN_ADD, x, y + i * (bh + gap), bw, bh, tag
    i = i + 1
    PlaceOne ws, BTN_DEL, x, y + i * (bh + gap), bw, bh, tag
    i = i + 1

    If runter Then
        PlaceOne ws, BTN_DOWN, x, y + i * (bh + gap), bw, bh, tag
    Else
        HideOne ws, BTN_DOWN, tag
    End If
End Sub


'  Naechste Planzeile oberhalb bzw. unterhalb, Ferienzeilen
'  uebersprungen. 0, wenn es dort keine mehr gibt.
Private Function InhaltszeileVor(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim i As Long
    For i = r - 1 To WP_FIRST_ROW Step -1
        If Not IsFerienRow(ws, i) Then
            InhaltszeileVor = i
            Exit Function
        End If
    Next i
End Function


Private Function InhaltszeileNach(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim i As Long, letzte As Long
    letzte = PlanLastRow(ws)
    For i = r + 1 To letzte
        If Not IsFerienRow(ws, i) Then
            InhaltszeileNach = i
            Exit Function
        End If
    Next i
End Function


'  Eine Schaltflaeche ausblenden - den Zeilenbereich aber TROTZDEM
'  nachfuehren.
'
'  ReadBlock holt sich den Bereich aus dem AlternativeText von
'  BTN_UP. Bliebe dort beim Ausblenden der alte Wert stehen, wuerden
'  "+" und "-" auf den Zeilen der vorherigen Auswahl arbeiten - also
'  an der voellig falschen Stelle einfuegen oder loeschen.
Private Sub HideOne(ByVal ws As Worksheet, ByVal nm As String, ByVal tag As String)
    Dim s As Shape
    On Error Resume Next
    Set s = Nothing
    Set s = ws.Shapes(nm)
    If Not s Is Nothing Then
        s.AlternativeText = tag
        s.Visible = msoFalse
    End If
    On Error GoTo 0
End Sub


Private Sub PlaceOne(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, _
                     ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                     ByVal tag As String)
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(nm)
    If s Is Nothing Then Exit Sub
    With s
        .LockAspectRatio = msoFalse
        .Left = x
        .Top = y
        .Width = w
        .Height = h
        .AlternativeText = tag
        .Visible = msoTrue
        .ZOrder msoBringToFront
    End With
End Sub


Private Sub HideButtons(ByVal ws As Worksheet)
    Dim nm As Variant, s As Shape
    For Each nm In Array(BTN_UP, BTN_ADD, BTN_DEL, BTN_DOWN)
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes(CStr(nm))
        On Error GoTo 0
        If Not s Is Nothing Then s.Visible = msoFalse
    Next nm
End Sub


Private Function ReadBlock(ByRef r1 As Long, ByRef r2 As Long) As Boolean
    Dim ws As Worksheet, s As Shape, t As String, p As Long
    Set ws = WpSheet()
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set s = ws.Shapes(BTN_UP)
    On Error GoTo 0
    If Not s Is Nothing Then
        t = s.AlternativeText
        p = InStr(t, ":")
        If p > 0 Then
            r1 = CLng(Left$(t, p - 1))
            r2 = CLng(Mid$(t, p + 1))
            ReadBlock = (r1 >= WP_FIRST_ROW And r2 >= r1)
        End If
    End If

    If Not ReadBlock Then
        If TypeName(Selection) = "Range" Then
            ReadBlock = GetBlock(ws, Selection, PlanLastRow(ws), r1, r2)
        End If
    End If
End Function


'=====================================================================
'  Hilfsfunktionen
'=====================================================================
Public Function WpSheet() As Worksheet
    On Error Resume Next
    Set WpSheet = ThisWorkbook.Worksheets(WP_SHEET)
End Function


Public Function SetSheet() As Worksheet
    On Error Resume Next
    Set SetSheet = ThisWorkbook.Worksheets(SET_SHEET)
End Function


'---------------------------------------------------------------------
'  Voellig leere Zeilen am unteren Ende des Plans entfernen.
'
'  Sie entstehen, sobald der Plan mehr Zeilen hat als es verfuegbare
'  Unterrichtswochen gibt - etwa wenn nachtraeglich Wochen abgewaehlt
'  werden oder eine Vorlage mehr Zeilen mitbringt als gebraucht.
'  Uebrig bleibt sonst ein Stueck leere Tabelle unter dem letzten
'  Eintrag.
'
'  Geloescht wird NUR, was wirklich nichts enthaelt. Eine Zeile mit
'  Thema oder Lehrplan-Code bleibt stehen, auch wenn ihr die
'  Unterrichtswoche fehlt - das ist Inhalt des Nutzers.
'---------------------------------------------------------------------
Public Function LeereEndzeilenEntfernen(ByVal ws As Worksheet) As Long
    Dim r As Long, letzte As Long, unten As Long, n As Long

    If ws Is Nothing Then Exit Function
    letzte = PlanLastRow(ws)

    '  ZUERST alles unterhalb der letzten Inhaltszeile.
    '
    '  Genau hier lag der Fehler, an dem nach dem ersten Import 35
    '  Leerzeilen stehengeblieben sind: die Schleife weiter unten
    '  beginnt bei PlanLastRow - und PlanLastRow sucht mit End(xlUp)
    '  nach INHALT. Zeilen, die nur formatiert sind (die Reserve, die
    '  die Vorlage mitbringt und die beim Einfuegen nach unten
    '  gerutscht ist), sieht es gar nicht. Sie lagen unterhalb dessen,
    '  was die Aufraeumschleife ueberhaupt betrachtet hat.
    '
    '  Unterhalb von PlanLastRow kann per Definition nichts mehr in
    '  E bis M, im Kennzeichen oder in den Hilfsspalten stehen - sonst
    '  haette PlanLastRow tiefer gezeigt. Der Block darf also am Stueck
    '  weg; das ist ausserdem deutlich schneller als zeilenweise.
    unten = UnterkanteBlatt(ws)
    If unten > letzte + 500 Then unten = letzte + 500      ' Notbremse
    If unten > letzte Then
        ws.Rows((letzte + 1) & ":" & unten).Delete Shift:=xlUp
        n = n + (unten - letzte)
    End If

    ' Dann von unten her die leeren Zeilen INNERHALB des Plans.
    ' Die erste Planzeile bleibt immer stehen.
    For r = letzte To WP_FIRST_ROW + 1 Step -1
        If IsFerienRow(ws, r) Then Exit For
        If Not ZeileOhneInhalt(ws, r) Then Exit For
        ws.Rows(r).Delete Shift:=xlUp
        n = n + 1
    Next r

    LeereEndzeilenEntfernen = n
End Function


'  Letzte Zeile, die Excel im Blatt ueberhaupt noch kennt - also auch
'  leere, aber formatierte Zeilen. UsedRange ist dafuer das einzige
'  brauchbare Mass; End(xlUp) taugt nicht, weil es nur Inhalt findet.
Private Function UnterkanteBlatt(ByVal ws As Worksheet) As Long
    Dim r As Long
    On Error Resume Next
    r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    On Error GoTo 0
    If r < WP_FIRST_ROW Then r = WP_FIRST_ROW
    UnterkanteBlatt = r
End Function


'  Das Haekchen in J zaehlt bewusst NICHT als Inhalt: es steht in
'  jeder Zeile und waere sonst immer gesetzt.
Private Function ZeileOhneInhalt(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array("E", "F", "G", "H", "I", "K", "M")
        If Len(Trim$(CStr(ws.Cells(r, CStr(c)).Value))) > 0 Then Exit Function
    Next c
    ZeileOhneInhalt = True
End Function


'=====================================================================
'  L E R N B E R E I C H E  -  Rechenspalten und Summenzeile
'  ------------------------------------------------------------------
'  Die Spalten F bis I und die Summenzeile sind reine Rechenspalten.
'  Bisher standen sie als feste Formeln im Blatt - mit fest
'  verdrahteten Bezuegen wie Einstellungen!$B$2 und $G$15:$G$52.
'  Damit war das Blatt "Lernbereiche" die einzige Stelle im Projekt,
'  die kaputtgeht, sobald jemand in "Einstellungen" eine Zeile
'  einfuegt. Ausserdem hat die Vorlage die Summenformel als
'  =SUM(E10:E10) mitgebracht, also ins Leere gezeigt.
'
'  Hier werden beide aus dem erkannten Layout neu gebaut. Was der
'  Anwender eintippt (A bis E und J), wird nicht angefasst.
'=====================================================================
Public Function Lernbereiche_Aufbauen() As Long
    Dim ws As Worksheet
    Dim z1 As Long, z2 As Long, zSum As Long, r As Long

    Set ws = LbSheet()
    If ws Is Nothing Then Exit Function
    If Not LayoutReady() Then Exit Function

    z1 = LB_FIRST_ROW()
    zSum = LbSummeRow(ws)
    z2 = LbLastDataRow(ws, zSum)

    If z2 >= z1 Then
        For r = z1 To z2
            ws.Cells(r, "F").Formula = FrmLbF(r)
            ws.Cells(r, "G").Formula = FrmLbG(r)
            ws.Cells(r, "H").Formula = FrmLbH(r)
            ws.Cells(r, "I").Formula = FrmLbI(r)
        Next r
        Lernbereiche_Aufbauen = z2 - z1 + 1
    End If

    '  Die Summenzeile bekommt genau den Bereich der Datenzeilen.
    '  Steht keine da, wird auch keine erfunden.
    If zSum > 0 And z2 >= z1 Then
        ws.Cells(zSum, "E").Formula = "=SUM(E" & z1 & ":E" & z2 & ")"
        ws.Cells(zSum, "F").Formula = "=SUM(F" & z1 & ":F" & z2 & ")"
        ws.Cells(zSum, "G").Formula = "=SUM(G" & z1 & ":G" & z2 & ")"
        ws.Cells(zSum, "H").Formula = "=SUM(H" & z1 & ":H" & z2 & ")"
    End If
End Function


Public Function LbSheet() As Worksheet
    On Error Resume Next
    Set LbSheet = ThisWorkbook.Worksheets(LB_SHEET)
End Function


'  Zeile der Summenzeile (0 = keine da). Erkannt am Wort "Summe" in
'  der Spalte "Lernbereich".
Public Function LbSummeRow(ByVal ws As Worksheet) As Long
    Dim r As Long, ende As Long
    If ws Is Nothing Then Exit Function
    ende = LB_FIRST_ROW() + LB_FILTER_RESERVE + 20
    For r = LB_FIRST_ROW() To ende
        If StrComp(Trim$(CStr(ws.Cells(r, "D").Value)), LBL_LBSUMME, vbTextCompare) = 0 Then
            LbSummeRow = r
            Exit Function
        End If
    Next r
End Function


'  Letzte Zeile mit einem Lernbereich. Gezaehlt wird bis zur
'  Summenzeile; Leerzeilen dazwischen beenden die Tabelle NICHT,
'  sonst wuerde eine versehentlich leergeraeumte Zeile in der Mitte
'  alles darunter aus der Summe werfen.
Public Function LbLastDataRow(ByVal ws As Worksheet, ByVal zSum As Long) As Long
    Dim r As Long, ende As Long, letzte As Long
    If ws Is Nothing Then Exit Function
    If zSum > 0 Then
        ende = zSum - 1
    Else
        ende = LB_FIRST_ROW() + LB_FILTER_RESERVE
    End If
    letzte = LB_FIRST_ROW() - 1
    For r = LB_FIRST_ROW() To ende
        If Len(Trim$(CStr(ws.Cells(r, "A").Value))) > 0 _
           Or Len(Trim$(CStr(ws.Cells(r, "D").Value))) > 0 Then letzte = r
    Next r
    LbLastDataRow = letzte
End Function


'  Die vier Rechenspalten. Alle Bezuege kommen aus dem erkannten
'  Layout - keine feste Zeilennummer.
Private Function FrmLbF(ByVal r As Long) As String
    ' Wochen laut Lehrplan = Stunden / Stunden pro Woche, aufgerundet
    FrmLbF = "=ROUNDUP(E" & r & "/" & CellRef("B", mRowStdWoche) & ",0)"
End Function

Private Function FrmLbG(ByVal r As Long) As String
    ' Geplante Stunden = geplante Wochen * Stunden pro Woche
    FrmLbG = "=IF(H" & r & "="""","""",H" & r & "*" & CellRef("B", mRowStdWoche) & ")"
End Function

Private Function FrmLbH(ByVal r As Long) As String
    ' Geplante Wochen = verfuegbare Wochen bis "bis" minus die bis "von"-1
    FrmLbH = "=IF(OR(B" & r & "="""",C" & r & "=""""),""""," & _
             "INDEX(" & ColRef("G") & ",MATCH(C" & r & "," & ColRef("A") & ",0))" & _
             "-IF(B" & r & "=1,0,INDEX(" & ColRef("G") & ",MATCH(B" & r & "-1," & ColRef("A") & ",0))))"
End Function

Private Function FrmLbI(ByVal r As Long) As String
    ' Noch verfuegbare Wochen ab diesem Lernbereich
    FrmLbI = "=IF(OR(B" & r & "="""",C" & r & "=""""),""""," & _
             "IF(B" & r & "=1," & CellRef("B", mRowKlWochen) & "," & _
             "INDEX(" & ColRef("H") & ",MATCH(B" & r & "-1," & ColRef("A") & ",0)))-H" & r & ")"
End Function


Public Function PlanLastRow(ByVal ws As Worksheet) As Long
    Dim cols As Variant, c As Variant, r As Long, m As Long
    cols = Array("E", "F", "G", "H", "I", "K", "M", MARK_COL, "V")
    m = WP_FIRST_ROW
    For Each c In cols
        r = ws.Cells(ws.Rows.Count, CStr(c)).End(xlUp).Row
        If r > m Then m = r
    Next c
    If m < WP_FIRST_ROW Then m = WP_FIRST_ROW
    PlanLastRow = m
End Function


'  Enthaelt der markierte Bereich mindestens eine normale Planzeile?
Private Function HatInhaltszeile(ByVal ws As Worksheet, ByVal r1 As Long, _
                                 ByVal r2 As Long) As Boolean
    Dim r As Long
    For r = r1 To r2
        If Not IsFerienRow(ws, r) Then
            HatInhaltszeile = True
            Exit Function
        End If
    Next r
End Function


Private Function GetBlock(ByVal ws As Worksheet, ByVal Target As Range, _
                          ByVal lastRow As Long, ByRef r1 As Long, ByRef r2 As Long) As Boolean
    Dim isect As Range, ar As Range, lo As Long, hi As Long
    If Target Is Nothing Then Exit Function
    If Not Target.Parent Is ws Then Exit Function

    Set isect = Application.Intersect(Target, ws.Range("A" & WP_FIRST_ROW & ":V" & lastRow))
    If isect Is Nothing Then Exit Function

    lo = ws.Rows.Count: hi = 0
    For Each ar In isect.Areas
        If ar.Row < lo Then lo = ar.Row
        If ar.Row + ar.Rows.Count - 1 > hi Then hi = ar.Row + ar.Rows.Count - 1
    Next ar
    r1 = lo
    r2 = hi
    If r1 < WP_FIRST_ROW Then r1 = WP_FIRST_ROW
    If r2 > lastRow Then r2 = lastRow
    GetBlock = (r2 >= r1)
End Function


Private Sub SelectBlock(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    On Error Resume Next
    If ActiveSheet Is ws Then
        ws.Range(ws.Cells(r1, CF_COL_FIRST), ws.Cells(r2, CF_COL_LAST)).Select
    End If
End Sub


'=====================================================================
'  Meldungen - im stillen Modus (Selbsttest) unterdrueckt
'=====================================================================
Public Sub SetQuiet(ByVal b As Boolean)
    mQuiet = b
    If b Then
        mLastInfo = ""
        mLastError = ""
    End If
End Sub

Public Function IsQuiet() As Boolean
    IsQuiet = mQuiet
End Function

Public Function LastInfo() As String
    LastInfo = mLastInfo
End Function

Public Function LastError() As String
    LastError = mLastError
End Function

Public Sub ClearLastError()
    mLastError = ""
End Sub


'  Windows spielt zu JEDEM Meldungssymbol einen Systemklang ab - bei
'  jeder Bestaetigung ein "Pling". Ohne Symbol bleibt der Dialog
'  stumm. Die Symbole haben hier ohnehin nie etwas getragen, was nicht
'  auch im Text steht. Der uebergebene Stil wird deshalb kurz vor dem
'  Anzeigen von seinem Symbolanteil befreit; Schaltflaechen
'  (vbOKCancel, vbDefaultButton2 ...) bleiben unberuehrt, und die
'  Auswertung fuer mLastError laeuft VORHER, arbeitet also weiter.
Private Function OhneTon(ByVal style As Long) As Long
    OhneTon = style And Not (vbCritical Or vbQuestion Or vbExclamation Or vbInformation)
End Function


'  Rueckfrage. Im stillen Modus wird automatisch mit OK geantwortet.
Public Function Frage(ByVal txt As String, ByVal style As Long, _
                      Optional ByVal titel As String = "Stoffverteilungsplan") As Long
    If mQuiet Then
        mLastInfo = txt
        Frage = vbOK
    Else
        Frage = MsgBox(txt, OhneTon(style), titel)
    End If
End Function


'  Meldung. Im stillen Modus nur gemerkt, nicht angezeigt.
Public Sub Info(ByVal txt As String, Optional ByVal style As Long = vbInformation, _
                Optional ByVal titel As String = "Stoffverteilungsplan")
    mLastInfo = txt
    If (style And vbExclamation) = vbExclamation Or (style And vbCritical) = vbCritical Then
        mLastError = txt
    End If
    If Not mQuiet Then MsgBox txt, OhneTon(style), titel
End Sub


'=====================================================================
'  Nur fuer den Selbsttest: der Nachweis, dass Option Private Module
'  die Schaltflaechen nicht lahmlegt.
'
'  Shape.OnAction loest einen Makronamen im eigenen Projekt auf -
'  genau wie Application.Run. Microsoft schreibt zu Option Private
'  Module, die Public-Teile blieben "innerhalb des Projekts, das das
'  Modul enthaelt, weiterhin verfuegbar"; gesperrt wird nur der
'  Zugriff FREMDER Projekte. Eine Zusicherung fuer OnAction steht
'  dort aber nicht. Statt sich auf den Umkehrschluss zu verlassen,
'  probiert der Selbsttest es einfach aus: laeuft diese Prozedur
'  ueber Application.Run an, laufen auch die Knoepfe.
'  Sie tut absichtlich nichts weiter, damit die Probe folgenlos ist.
'=====================================================================
Public Sub SelbsttestProbe()
    mProbeGelaufen = True
End Sub


Public Function ProbeGelaufen() As Boolean
    ProbeGelaufen = mProbeGelaufen
    mProbeGelaufen = False               ' fuer den naechsten Lauf zuruecksetzen
End Function


Public Sub SetStep(ByVal s As String)
    mStep = s
End Sub


Public Function CurrentStep() As String
    CurrentStep = mStep
End Function


Public Sub ReportError(ByVal what As String)
    Dim n As Long, d As String, hinweis As String
    n = Err.Number
    d = Err.Description

    ' Haeufigste Ursache fuer Laufzeitfehler 1004 in diesem Projekt:
    ' ein Blatt liess sich nicht entsperren (falsches Kennwort in
    ' modKonfig). Das steht sonst nirgends und kostet sonst eine
    ' halbe Stunde Suche.
    If modSchutz.EntsperrFehler() <> "" Then
        hinweis = vbCrLf & vbCrLf & _
                  "Vermutliche Ursache: der Blattschutz liess sich nicht aufheben." & vbCrLf & _
                  "Betroffene Blätter: " & modSchutz.EntsperrFehler() & vbCrLf & _
                  "Bitte das Kennwort in SCHUTZ_PW (Modul modKonfig) prüfen."
    End If

    FastOff
    Info "Fehler beim " & what & "." & vbCrLf & vbCrLf & _
           "Schritt : " & mStep & vbCrLf & _
           "Nummer  : " & n & vbCrLf & _
           "Meldung : " & d & hinweis, vbExclamation, "Stoffverteilungsplan"
End Sub


'  onlySheet: das einzige Blatt, das die Aktion veraendert. Dann wird
'  auch nur dieses Blatt entsperrt und nur dieses neu berechnet - das
'  ist um ein Vielfaches schneller als der Durchlauf ueber alle
'  Blaetter. Ohne Angabe laeuft der volle Weg (fuer Aktionen, die
'  Blaetter anlegen oder mehrere Blaetter aendern).
Public Sub FastOn(Optional ByVal onlySheet As Worksheet = Nothing)
    If mBusy Then Exit Sub
    mBusy = True
    Set mScopeSheet = onlySheet
    With Application
        mScreenSave = .ScreenUpdating
        mEventSave = .EnableEvents
        mAlertSave = .DisplayAlerts
        mCalcSave = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    Set mActiveSave = Nothing
    On Error Resume Next
    Set mActiveSave = ActiveSheet
    On Error GoTo 0
    modSchutz.Schutz_Aus onlySheet
End Sub


Public Sub FastOff()
    If Not mBusy Then Exit Sub
    modSchutz.Schutz_An
    With Application
        .Calculation = mCalcSave
        If mScopeSheet Is Nothing Then
            .Calculate
        Else
            mScopeSheet.Calculate
        End If
        .DisplayAlerts = mAlertSave
        .EnableEvents = mEventSave
    End With
    ' Erst ALLES erledigen, was Blaetter anfasst - danach den Nutzer
    ' dorthin zuruecksetzen, wo er war. Umgekehrt hat es nicht
    ' funktioniert: SeitenumbruecheAus lief nach dem Zuruecksetzen und
    ' hat den Anwender nach "Blattschutz ein" auf der Anleitung
    ' abgesetzt. Das Zuruecksetzen ist jetzt die LETZTE Aktion.
    Set mScopeSheet = Nothing
    SeitenumbruecheAus

    ' Kein ungewolltes Blattwechseln: wo der Nutzer war, bleibt er
    ' auch. Gezielte Wechsel machen die Makros nach FastOff ueber
    ' GotoSheet.
    On Error Resume Next
    If Not mActiveSave Is Nothing Then
        If Not ActiveSheet Is mActiveSave Then mActiveSave.Activate
    End If
    On Error GoTo 0
    Set mActiveSave = Nothing
    Application.ScreenUpdating = mScreenSave
    mBusy = False
End Sub


'  Gezielter Blattwechsel nach einer Aktion - erst nach FastOff
'  aufrufen, sonst macht FastOff ihn wieder rueckgaengig.
'  Die gestrichelten Seitenumbruch-Linien blendet Excel nach jeder
'  Druckvorschau und nach jeder PageSetup-Aenderung wieder ein. Diese
'  Einstellung ueberlebt das Speichern nicht - deshalb wird sie nach
'  jeder Makroaktion und beim Oeffnen der Mappe neu gesetzt.
Public Sub SeitenumbruecheAus()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.DisplayPageBreaks Then ws.DisplayPageBreaks = False
    Next ws
End Sub


Public Sub GotoSheet(ByVal nm As String, Optional ByVal cellAddr As String = "")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    If cellAddr <> "" Then ws.Range(cellAddr).Select
End Sub


'=====================================================================
'  D I A G N O S E
'=====================================================================
Public Sub Diag_Wochenplan()
    Dim ws As Worksheet, rep As String, lastRow As Long, outPath As String
    Dim cr() As Long, nRows As Long, r As Long, nFer As Long
    Dim befund As String

    Set ws = WpSheet()
    If ws Is Nothing Then
        Info "Tabellenblatt '" & WP_SHEET & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    lastRow = PlanLastRow(ws)
    On Error GoTo 0
    If lastRow < WP_FIRST_ROW Then lastRow = 32

    Application.ScreenUpdating = False

    rep = "=== Diagnose Wochenplan ===" & vbCrLf
    rep = rep & "Zeitpunkt       : " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    rep = rep & "Excel-Version   : " & Application.Version & "  (Build " & Application.Build & ")" & vbCrLf
    rep = rep & "Sprache Oberfl. : " & Application.LanguageSettings.LanguageID(2) & vbCrLf
    rep = rep & "Listentrenner   : " & Application.International(xlListSeparator) & vbCrLf
    rep = rep & "Letzte Planzeile: " & lastRow & vbCrLf

    nRows = ContentRows(ws, lastRow, cr)
    For r = WP_FIRST_ROW To lastRow
        If IsFerienRow(ws, r) Then nFer = nFer + 1
    Next r
    rep = rep & "Inhaltszeilen   : " & nRows & vbCrLf
    rep = rep & "Ferienzeilen    : " & nFer & vbCrLf
    rep = rep & "Fixiert         : " & IIf(IsPlanFixed(), "ja", "nein") & vbCrLf
    rep = rep & "Blattschutz     : " & IIf(modSchutz.SchutzAktiv(), "aktiv", "aus") & vbCrLf
    If IsPlanFixed() Then
        rep = rep & "Naechste Nummer : " & NextRefNumber(ws, lastRow) & vbCrLf
    End If

    rep = rep & vbCrLf & "--- Erkanntes Layout 'Einstellungen' ---" & vbCrLf
    If RefreshLayout(True) Then
        rep = rep & LayoutInfo() & vbCrLf
    Else
        rep = rep & "Kalendertabelle NICHT gefunden!" & vbCrLf
    End If

    befund = modPruefung.MappePruefen()
    rep = rep & vbCrLf & "--- Analyse der Mappe ---" & vbCrLf
    If Len(befund) = 0 Then
        rep = rep & "keine Auffaelligkeiten" & vbCrLf
    Else
        rep = rep & befund
    End If

    rep = rep & vbCrLf & "--- Formen ---" & vbCrLf
    rep = rep & DiagShapes(ws)

    Application.ScreenUpdating = True

    rep = rep & vbCrLf & "=== Ende ===" & vbCrLf
    Debug.Print rep
    WriteReport rep, outPath

    Info "Diagnose abgeschlossen." & vbCrLf & vbCrLf & _
           "Bericht: " & outPath, vbInformation, "Stoffverteilungsplan"
End Sub


Private Function DiagShapes(ByVal ws As Worksheet) As String
    Dim s As String, nm As Variant, sh As Shape
    On Error Resume Next
    Err.Clear
    EnsureButtons ws
    If Err.Number <> 0 Then
        s = s & "EnsureButtons: FEHLER " & Err.Number & " - " & Err.Description & vbCrLf
        Err.Clear
    Else
        s = s & "EnsureButtons: ok" & vbCrLf
    End If
    For Each nm In Array(BTN_UP, BTN_ADD, BTN_DEL, BTN_DOWN)
        Set sh = Nothing
        Set sh = ws.Shapes(CStr(nm))
        If sh Is Nothing Then
            s = s & "  " & nm & ": fehlt" & vbCrLf
        Else
            s = s & "  " & nm & ": AutoShapeType " & sh.AutoShapeType & _
                    ", Makro " & sh.OnAction & vbCrLf
        End If
        Err.Clear
    Next nm
    HideButtons ws
    DiagShapes = s
End Function


Private Sub WriteReport(ByVal txt As String, ByRef outPath As String)
    Dim p As String, ff As Integer
    p = ThisWorkbook.Path
    If p = "" Or LCase$(Left$(p & "    ", 4)) = "http" Then p = DIAG_FALLBACK_DIR
    outPath = p & "\Wochenplan_Diagnose.txt"

    On Error Resume Next
    Err.Clear
    ff = FreeFile
    Open outPath For Output As #ff
    Print #ff, txt
    Close #ff
    If Err.Number <> 0 Then
        Err.Clear
        outPath = DIAG_FALLBACK_DIR & "\Wochenplan_Diagnose.txt"
        ff = FreeFile
        Open outPath For Output As #ff
        Print #ff, txt
        Close #ff
        If Err.Number <> 0 Then outPath = "(Datei konnte nicht geschrieben werden)"
    End If
End Sub
