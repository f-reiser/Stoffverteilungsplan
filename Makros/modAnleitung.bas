Attribute VB_Name = "modAnleitung"
Option Explicit

'  Option Private Module: alles, was in diesem Modul Public ist, bleibt
'  fuer die anderen Module dieses Projekts voll erreichbar - es
'  verschwindet nur aus der Makroliste (Alt+F8) und aus dem Zugriff
'  FREMDER VBA-Projekte. Genau das ist hier gewollt: die Liste hat
'  zuletzt 30 Eintraege gehabt, von denen fuenf gemeint waren.
'  Was von Hand gestartet werden soll, steht in modStart.
Option Private Module

'=====================================================================
'  Stoffverteilungsplan - das Anleitungsblatt
'  ------------------------------------------------------------------
'    "Anleitung"   fuer alle, die die Mappe fertig eingerichtet
'                  bekommen - der Weg zum Plan in drei Schritten
'
'  Anleitungsblaetter_Sicherstellen legt nur an, was fehlt, und
'  laesst ein vorhandenes Blatt in Ruhe - von Hand gemachte
'  Aenderungen bleiben also erhalten.
'
'  Anleitungsblaetter_Erzeugen schreibt das Blatt komplett neu.
'  Bewusst OHNE Schaltflaeche: der Aufruf geht ueber Alt+F8 (Eintrag
'  "Anleitungsblatt_neu_schreiben" in modStart) und ueberschreibt
'  alles, was jemand von Hand ergaenzt hat.
'
'  ENTFALLEN ist das zweite Blatt "VBA-Update" bzw. "Update". Es hat
'  erklaert, wie man Module von Hand importiert - und spaeter, wie
'  die Uebernahme funktioniert. Beides steht jetzt dort, wo man es
'  wirklich liest: der Update-Weg als Erklaertext neben der
'  Schaltflaeche "Daten importieren" im Blatt "Steuerung", die
'  Entwicklerseite in Makros\LIESMICH.txt. Ein eigenes Blatt dafuer
'  war fuer Kollegen nur ein Reiter mehr, den niemand anfasst.
'  AltesBlattEntfernen raeumt es beim naechsten Einrichten weg.
'=====================================================================

Public Const SHEET_HELP As String = "Anleitung"

'  Stand des Anleitungstextes. Bei jeder inhaltlichen Aenderung an
'  BuildHelpSheet HOCHZAEHLEN.
'
'  Warum ueberhaupt: Anleitungsblaetter_Sicherstellen hat bisher nur
'  angelegt, was FEHLT. Ein vorhandenes Blatt blieb stehen - und damit
'  auch sein alter Text. Der Abschnitt "Eine neue Fassung der Mappe"
'  ist deshalb nach dem Update nirgends aufgetaucht, obwohl er im Code
'  stand. Jetzt wird das Blatt neu geschrieben, sobald sein Stand
'  aelter ist als dieser hier.
'  Preis dafuer: von Hand ergaenzter Text im Blatt "Anleitung" geht
'  bei einem Versionswechsel verloren. Das ist bewusst so - das Blatt
'  ist erzeugter Text, kein Notizzettel; ein veralteter Text kostet
'  mehr als eine verlorene Randnotiz.
Public Const ANLEITUNG_STAND As String = "2026-09-05a"

'  Unsichtbarer Name in der Mappe, in dem der Stand hinterlegt wird.
'  Public, damit der Mutationstest ihn gezielt verstellen kann.
Public Const STAND_NAME As String = "wpAnleitungStand"

'  Die beiden Namen, unter denen das entfallene Blatt in aelteren
'  Mappen steckt.
Private Const SHEET_ALT_1 As String = "VBA-Update"
Private Const SHEET_ALT_2 As String = "Update"

Private Const TXT_COL   As String = "B"
Private Const TXT_WIDTH As Double = 118

Private mRow As Long


'=====================================================================
'  Oeffentliche Makros
'=====================================================================

'  Legt nur an, was fehlt. Rueckgabe True, wenn etwas entstanden ist.
Public Function Anleitungsblaetter_Sicherstellen() As Boolean
    Dim fehlt As Boolean, veraltet As Boolean

    '  Rueckgabe der Funktion = "es hat sich etwas geaendert". Ein
    '  geloeschtes Update-Blatt zaehlt dazu.
    Anleitungsblaetter_Sicherstellen = AltesBlattEntfernen()

    fehlt = (SheetOrNothing(SHEET_HELP) Is Nothing)
    veraltet = (AnleitungStand() <> ANLEITUNG_STAND)
    If Not fehlt And Not veraltet Then Exit Function

    modWochenplan.FastOn
    On Error GoTo Fail
    modWochenplan.SetStep "Blatt " & SHEET_HELP
    BuildHelpSheet PrepareSheet(SHEET_HELP)
    modWochenplan.FastOff
    Anleitungsblaetter_Sicherstellen = True
    Exit Function
Fail:
    modWochenplan.ReportError "Anlegen der Anleitungsblätter"
End Function


'  Der Stand, der in dieser Mappe hinterlegt ist ("" = keiner).
'  Hinterlegt wird er als unsichtbarer Name, nicht als Zellinhalt:
'  eine Zelle koennte jemand ueberschreiben, und im Blattbild haette
'  sie nichts zu suchen.
Public Function AnleitungStand() As String
    Dim s As String
    On Error Resume Next
    s = CStr(ThisWorkbook.Names(STAND_NAME).RefersTo)
    On Error GoTo 0
    s = Replace(s, "=", "")
    s = Replace(s, """", "")
    AnleitungStand = Trim$(s)
End Function


Private Sub AnleitungStandSetzen()
    On Error Resume Next
    ThisWorkbook.Names(STAND_NAME).Delete
    ThisWorkbook.Names.Add Name:=STAND_NAME, _
                           RefersTo:="=""" & ANLEITUNG_STAND & """", _
                           Visible:=False
    On Error GoTo 0
End Sub


'  Das entfallene Update-Blatt aus aelteren Mappen wegraeumen.
'
'  ZWEI STOLPERSTEINE, beide beim ersten Anlauf zugeschlagen:
'
'  1. Die Mappe ist STRUKTURGESCHUETZT (Blattschutz_Einrichten setzt
'     ThisWorkbook.Protect Structure:=True). Dagegen laeuft jedes
'     Worksheets(...).Delete ins Leere - und unter On Error Resume
'     Next voellig lautlos. Das Blatt blieb einfach stehen, ohne
'     Fehlermeldung. Der Schutz wird deshalb hier gezielt aufgehoben
'     und danach wieder gesetzt.
'  2. Excel fragt beim Loeschen eines Blattes nach, wenn DisplayAlerts
'     an ist. Innerhalb von FastOn waere es ohnehin aus, aber diese
'     Routine laeuft auch ausserhalb.
'
'  Rueckgabe: True, wenn wirklich etwas geloescht wurde.
Private Function AltesBlattEntfernen() As Boolean
    Dim nm As Variant, ws As Worksheet
    Dim alertSave As Boolean, warGeschuetzt As Boolean

    On Error Resume Next
    alertSave = Application.DisplayAlerts
    Application.DisplayAlerts = False

    warGeschuetzt = ThisWorkbook.ProtectStructure
    If warGeschuetzt Then ThisWorkbook.Unprotect SCHUTZ_PW

    For Each nm In Array(SHEET_ALT_1, SHEET_ALT_2)
        Set ws = SheetOrNothing(CStr(nm))
        If Not ws Is Nothing Then
            '  Ein ausgeblendetes Blatt laesst sich nicht loeschen,
            '  und die Mappe braucht mindestens ein sichtbares.
            ws.Visible = xlSheetVisible
            ws.Delete
            If SheetOrNothing(CStr(nm)) Is Nothing Then AltesBlattEntfernen = True
        End If
        Set ws = Nothing
    Next nm

    If warGeschuetzt And Not ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Protect Password:=SCHUTZ_PW, Structure:=True, Windows:=False
    End If
    Application.DisplayAlerts = alertSave
    On Error GoTo 0
End Function


'  Schreibt beide Blaetter neu - ueberschreibt eigene Ergaenzungen!
Public Sub Anleitungsblaetter_Erzeugen()
    If modWochenplan.Frage("Das Blatt """ & SHEET_HELP & """ wird komplett neu " & _
              "geschrieben." & vbCrLf & vbCrLf & _
              "Alles, was dort von Hand ergänzt oder geändert wurde, geht dabei " & _
              "verloren. Fortfahren?", _
              vbOKCancel + vbDefaultButton2, _
              "Stoffverteilungsplan") <> vbOK Then Exit Sub

    modWochenplan.FastOn
    On Error GoTo Fail

    AltesBlattEntfernen
    modWochenplan.SetStep "Blatt " & SHEET_HELP
    BuildHelpSheet PrepareSheet(SHEET_HELP)

    modWochenplan.FastOff
    modSchutz.Blattschutz_Einrichten
    Exit Sub
Fail:
    modWochenplan.ReportError "Aufbauen der Anleitungsblätter"
End Sub


'=====================================================================
'  Blattgeruest
'=====================================================================
Private Function PrepareSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet

    Set ws = SheetOrNothing(nm)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If

    ws.Cells.Clear
    ws.Cells.UnMerge
    ws.Cells.Interior.Pattern = xlNone

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 3
    ws.Columns(TXT_COL).ColumnWidth = TXT_WIDTH
    ws.Rows(1).RowHeight = 10

    mRow = 2
    Set PrepareSheet = ws
End Function


Private Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(nm)
End Function


'---------------------------------------------------------------------
'  Eine Zeile schreiben. style:
'    H1  Blatttitel      H2  Abschnitt      H3  Zwischenzeile
'    P   Fliesstext      S   Aufzaehlung
'    C   Code / Klickpfad     N  Hinweiskasten     -  Leerzeile
'---------------------------------------------------------------------
Private Sub W(ByVal ws As Worksheet, ByVal style As String, ByVal txt As String)
    Dim c As Range

    Set c = ws.Cells(mRow, TXT_COL)

    Select Case UCase$(style)
        Case "-"
            ws.Rows(mRow).RowHeight = 9
            mRow = mRow + 1
            Exit Sub

        Case "H1"
            c.Value = txt
            c.Font.Size = 22
            c.Font.Bold = True
            c.Font.Color = RGB(45, 55, 72)
            ws.Rows(mRow).RowHeight = 32

        Case "H2"
            c.Value = txt
            c.Font.Size = 13
            c.Font.Bold = True
            c.Font.Color = RGB(255, 255, 255)
            c.Interior.Pattern = xlSolid
            c.Interior.Color = modWochenplan.FARBE_BALKEN
            c.IndentLevel = 1
            c.VerticalAlignment = xlCenter
            ws.Rows(mRow).RowHeight = 26

        Case "H3"
            c.Value = txt
            c.Font.Size = 11
            c.Font.Bold = True
            c.Font.Color = RGB(60, 72, 92)
            ws.Rows(mRow).RowHeight = 22

        Case "C"
            c.Value = txt
            c.Font.Name = "Consolas"
            c.Font.Size = 10
            c.Font.Color = RGB(40, 60, 90)
            c.Interior.Pattern = xlSolid
            c.Interior.Color = RGB(240, 243, 247)
            c.IndentLevel = 2
            c.WrapText = True
            c.VerticalAlignment = xlCenter
            ws.Rows(mRow).AutoFit

        Case "N"
            c.Value = txt
            c.Font.Size = 10.5
            c.Font.Italic = True
            c.Font.Color = RGB(120, 85, 20)
            c.Interior.Pattern = xlSolid
            c.Interior.Color = RGB(253, 246, 227)
            c.IndentLevel = 1
            c.WrapText = True
            c.VerticalAlignment = xlCenter
            ws.Rows(mRow).AutoFit

        Case "S"
            c.Value = txt
            c.Font.Size = 11
            c.Font.Color = RGB(35, 45, 60)
            c.IndentLevel = 2
            c.WrapText = True
            c.VerticalAlignment = xlTop
            ws.Rows(mRow).AutoFit

        Case Else                                   ' "P"
            c.Value = txt
            c.Font.Size = 11
            c.Font.Color = RGB(35, 45, 60)
            c.IndentLevel = 1
            c.WrapText = True
            c.VerticalAlignment = xlTop
            ws.Rows(mRow).AutoFit
    End Select

    mRow = mRow + 1
End Sub


'=====================================================================
'  Blatt 1 - der Weg zum Plan
'=====================================================================
Private Sub BuildHelpSheet(ByVal ws As Worksheet)

    W ws, "H1", "Stoffverteilungsplan - Kurzanleitung"
    W ws, "P", "Eine Mappe je Klasse und Fach. Du arbeitest dich in drei Schritten " & _
               "vor: erst klärst du, welche Wochen das Schuljahr überhaupt hergibt, " & _
               "dann verteilst du die Lernbereiche grob darauf, und zum Schluss " & _
               "planst du Woche für Woche. Das Rechnen und Prüfen übernimmt die Mappe."
    W ws, "N", "Die Makros sind bereits eingebaut. Du musst nichts importieren und " & _
               "nichts über Alt+F8 starten - alles läuft über die Schaltflächen."
    W ws, "-", ""

    ' ---------------- Schritt 1 --------------------------------------
    W ws, "H2", "Schritt 1   Einstellungen - welche Wochen hast du überhaupt?"
    W ws, "P", "Ziel: Am Ende dieses Schritts weißt du, wie viele Unterrichtswochen " & _
               "dir in dieser Klasse wirklich bleiben."
    W ws, "S", "Oben links Schuljahresstart (das Jahr, in dem das Schuljahr beginnt) " & _
               "und Stunden pro Woche eintragen."
    W ws, "S", "Im Schulwochen-Kalender bei jeder Woche das Häkchen ""Woche verfügbar"" " & _
               "setzen oder wegnehmen. Wandertage, Praktika und Prüfungswochen also " & _
               "abwählen. ""halbe Woche"" und ""halbe Klasse"" für alles, was nur zur " & _
               "Hälfte stattfindet - beides zählt in der Stundenrechnung mit."
    W ws, "S", "Ferientabelle rechts daneben: die Schaltfläche ""Ferien aus Kalender " & _
               "vorschlagen"" darüber füllt sie aus den Lücken zwischen den Schulwochen. " & _
               "Bezeichnungen kurz prüfen, fertig. Die Spalte heißt bewusst """ & _
               LBL_FERNAME & """: dort gehört alles hinein, was Unterricht kostet - " & _
               "auch Fortbildungen, Praktika oder ein Schullandheimaufenthalt."
    W ws, "S", "Die vier Kopfangaben eintragen: " & LBL_FACH & ", " & LBL_KLASSE & _
               ", " & LBL_SCHULE & " (Auswahlliste) und " & LBL_LEHRER & ". Danach " & _
               """Kopf übernehmen"" drücken - das schreibt über die Tabellen in """ & _
               WP_SHEET & """ und """ & LB_SHEET & """ eine Titelzeile: links das " & _
               "Schullogo, daneben Fach und Klasse, rechts Lehrkraft und Schuljahr."
    W ws, "N", "Die Titelzeile braucht nur eine einzige Zeile. Der Schulname steht " & _
               "nicht extra dabei - er steht schon im Logo. Beide Schullogos stecken " & _
               "fertig in der Mappe; welches erscheint, entscheidet allein die " & _
               "Auswahl bei """ & LBL_SCHULE & """."
    W ws, "N", "Die Zahl, auf die es ankommt, steht oben links: ""verfügbare " & _
               "Klassen-Wochen"". Mit ihr planst du im nächsten Schritt."
    W ws, "-", ""

    ' ---------------- Schritt 2 --------------------------------------
    W ws, "H2", "Schritt 2   Lernbereiche - die Grobplanung"
    W ws, "P", "Ziel: Du verteilst die Lernbereiche des LehrplanPLUS auf das " & _
               "Schuljahr - noch ohne einzelne Stunden."
    W ws, "S", "Je Lernbereich eine Zeile: Referenzcode (z. B. M10.1), Bezeichnung " & _
               "und ""Std. lt. Lehrplan"". Die Kompetenzerwartungen kannst du rechts " & _
               "in der letzten Spalte sammeln."
    W ws, "S", "Dann ""von"" und ""bis"" als Unterrichtswoche eintragen. Die Spalten " & _
               "rechts rechnen sofort mit: wie viele Wochen der Lehrplan vorsieht, " & _
               "wie viele du tatsächlich eingeplant hast und was bis zum Jahresende " & _
               "noch übrig bleibt."
    W ws, "S", "So lange schieben, bis die Summe unten zum Erwartungswert des " & _
               "Lehrplans passt und am Jahresende noch Luft für Puffer bleibt."
    W ws, "N", "Der Referenzcode ist das Bindeglied zwischen den Blättern: Er taucht " & _
               "im Wochenplan in Spalte F wieder auf. Passt ein Code dort nicht zu " & _
               "der Woche, in der die Zeile steht, meldet sich die Warnspalte. " & _
               "Prüfungszeilen sind ausgenommen - sie enden auf "".P"" (z. B. M10.P)."
    W ws, "-", ""

    ' ---------------- Schritt 3 --------------------------------------
    W ws, "H2", "Schritt 3   Wochenplan - die Feinplanung"
    W ws, "P", "Ziel: Jede Unterrichtswoche bekommt ihr Thema."
    W ws, "S", "Einmal im Blatt """ & CTRL_SHEET & """ auf ""Wochenplan neu " & _
               "aufbauen"" drücken. Danach hat jede Zeile ihre Unterrichtswoche " & _
               "in Spalte E, und dazwischen stehen die grauen Ferienzeilen."
    W ws, "N", "Diesen Knopf darfst du jederzeit wieder drücken, auch wenn der Plan " & _
               "schon voll ist. Er ordnet nur die Unterrichtswochen und die " & _
               "Ferienzeilen neu - deine Einträge in den Spalten F bis M bleiben " & _
               "unangetastet und wandern mit ihrer Zeile mit."
    W ws, "S", "Jetzt Zeile für Zeile füllen: Lehrplan-Code (F), Thema (G), " & _
               "Kompetenzen (H), Material (I), Art der Stunde (K) und bei Bedarf " & _
               "Notizen (M)."
    W ws, "S", "Die Auswahl in Spalte K färbt die Zeile ein. So siehst du auf einen " & _
               "Blick, wo Schulaufgaben, Puffer und verschiebbare Stunden liegen."
    W ws, "S", "Reihenfolge ändern: Zeilen markieren - links am Rand erscheinen vier " & _
               "kleine Schaltflächen (hoch, plus, minus, runter). Die Unterrichtswochen " & _
               "bleiben stehen, es wandern nur die Inhalte, und Ferienzeilen werden " & _
               "übersprungen."
    W ws, "N", "Steht oben rechts ""Warnungen"", passt irgendwo etwas nicht zusammen. " & _
               "Die betroffene Zeile hat in Spalte A ein ""!""."
    W ws, "-", ""

    ' ---------------- Danach -----------------------------------------
    W ws, "H2", "Wenn die Planung steht"
    W ws, "P", "Solange der Plan offen ist, wird der Referenz Code in Spalte B " & _
               "berechnet - er ändert sich also, sobald Zeilen verschoben oder " & _
               "eingefügt werden."
    W ws, "P", "Sobald du die Codes verwendest bzw. weitergibst (z. B. in deinem " & _
               "Lehrerplaner oder weil du dich mit einem Kollegen über den Code " & _
               "abstimmst), klick im Blatt """ & CTRL_SHEET & """ auf " & _
               """Stoffverteilungsplan fixieren"". Danach sind die Codes feste Werte: " & _
               "sie wandern beim Verschieben mit, bleiben beim Löschen anderer Zeilen " & _
               "unverändert, und eine neue Zeile bekommt einen noch nie vergebenen " & _
               "Code, sobald du in Spalte F den Lehrplan-Code einträgst."
    W ws, "N", "''Fixierung aufheben'' nummeriert ALLE Codes neu durch. Nur benutzen, " & _
               "solange noch niemand mit den alten Codes arbeitet."
    W ws, "-", ""

    W ws, "H2", "Als PDF weitergeben"
    W ws, "S", "Im Blatt """ & CTRL_SHEET & """ unter der Schaltfläche das " & _
               "Papierformat wählen (DIN A4 oder - bei vielen Spalten angenehmer " & _
               "zu lesen - DIN A3), dann auf ""Als PDF exportieren"" klicken."
    W ws, "S", "Das PDF enthält zuerst den kompletten """ & WP_SHEET & """, danach die " & _
               """" & LB_SHEET & """, in einer einzigen Datei."
    W ws, "S", "Gedruckt wird immer im Querformat, und die Breite wird so skaliert, dass " & _
               "die Spalten " & PDF_WP_SPALTE_ERSTE & " bis " & PDF_WP_SPALTE_LETZTE & _
               " auf die Seite " & _
               "passen. Es rutscht also nichts seitlich auf eine zweite Seite."
    W ws, "S", "Es kommt der gewohnte Speichern-Dialog: Ordner und Dateiname sind " & _
               "frei wählbar, vorgeschlagen wird der Ordner der Mappe und ein Name " & _
               "aus Fach und Klasse."
    W ws, "N", "Der Titelblock erscheint nur einmal ganz oben, nicht auf jeder Seite - " & _
               "das spart Platz. Wiederholt wird auf jeder Seite die Überschriftenzeile."
    W ws, "-", ""

    W ws, "H2", "Im laufenden Schuljahr"
    W ws, "S", "Erledigtes in Spalte J abhaken - die Zeile wird dann leicht grün " & _
               "und unterstrichen dargestellt."
    W ws, "S", "Was verschoben werden muss, mit den vier Schaltflächen umsortieren."
    W ws, "S", "Ändert sich der Kalender (zusätzlicher Feiertag, ausgefallene Woche), " & _
               "das Häkchen in den Einstellungen anpassen und einmal " & _
               """Wochenplan neu aufbauen"" drücken."
    W ws, "-", ""

    ' ---------------- Nachschlagen -----------------------------------
    W ws, "H2", "Zum Nachschlagen"

    W ws, "H3", "Die Spalten im Blatt " & WP_SHEET
    W ws, "S", "A  Warnkennzeichen - ein ""!"" heißt: in dieser Zeile stimmt etwas nicht."
    W ws, "S", "B  Referenz Code - berechnet oder nach dem Fixieren ein fester Wert."
    W ws, "S", "C  Datum, D  KW - kommen automatisch aus der Unterrichtswoche."
    W ws, "S", "E  UW - die Unterrichtswoche. Wird erzeugt, bleibt aber von Hand änderbar."
    W ws, "S", "F  Lehrplan-Code - z. B. M10.1, bei Prüfungen M10.P."
    W ws, "S", "G  Thema     H  Kompetenzen     I  Material / Aufgaben"
    W ws, "S", "J  Erledigt - Kontrollkästchen. Abgehakte Zeilen werden farblich " & _
               "hervorgehoben."
    W ws, "S", "K  Stunde - Auswahlliste: Stofferarbeitung, Vorbereitung SA, " & _
               "SA/KA/kasL, Puffer, Verschiebbar. Je nach Auswahl bekommt " & _
               "die Zeile eine andere Farbe."
    W ws, "S", "L  Hinweis - zeigt automatisch ""1/2 Woche"" bzw. ""1/2 Klasse""."
    W ws, "S", "M  Notizen - freier Text."
    W ws, "S", "N bis V sind Hilfsspalten der Automatik. Bitte nicht von Hand ändern. " & _
               "In Spalte N steht bei den grauen Zwischenzeilen ein Kennzeichen - " & _
               "in der Hintergrundfarbe geschrieben, also nicht zu sehen. Die " & _
               "Einzelheiten zu den Warnungen stehen in den ausgeblendeten Spalten " & _
               "P bis S - über das kleine Plus oberhalb der Spaltenköpfe einblendbar."

    W ws, "H3", "Was gesperrt ist und warum"
    W ws, "S", "Beschreibbar sind nur die Zellen, die du wirklich pflegst: im " & _
               WP_SHEET & " die Spalten E bis K und M, in " & LB_SHEET & " die " & _
               "Eingabespalten, in " & SET_SHEET & " die Kopfwerte, der Kalender, " & _
               "die Ferientabelle und die vier Kopfangaben."
    W ws, "S", "Im " & WP_SHEET & " lassen sich Zeilen über das Excel-Kontextmenü " & _
               "weder einfügen noch löschen, verschieben oder sortieren. Der Plan " & _
               "lebt davon, dass Zeilen, Unterrichtswochen, Ferienzeilen und Formeln " & _
               "zusammenpassen - dafür gibt es die vier Schaltflächen am Zeilenrand."
    W ws, "S", "In " & LB_SHEET & " und " & SET_SHEET & " darfst du Zeilen einfügen " & _
               "und löschen. Unterhalb der Daten stehen ausserdem freie Zeilen bereit."
    W ws, "S", "Die Stundenliste in " & SET_SHEET & " ist gesperrt: an ihr hängen die " & _
               "Farbregeln und die Auswahlliste der Spalte K."
    W ws, "S", "Musst du doch einmal ganz frei arbeiten: """ & CTRL_SHEET & """, " & _
               """Blattschutz ein / aus"" - und danach bitte wieder einschalten."

    W ws, "H3", "Eine neue Fassung der Mappe"
    W ws, "S", "Neue Fassungen kommen als fertige, leere Datei - sie bringen die " & _
               "aktuellen Makros schon mit. Du importierst nichts und stellst nichts um."
    W ws, "S", "Die neue Datei öffnen, Makros zulassen, im Blatt """ & CTRL_SHEET & _
               """ unter ""Update"" auf ""Daten importieren"" klicken und deine " & _
               "bisherige Datei auswählen."
    W ws, "S", "Danach speichern - am besten unter dem Namen der bisherigen Datei. " & _
               "Deine alte Datei wird beim Import nur gelesen und bleibt unverändert."

    W ws, "P", "Übernommen wird alles, was du selbst eingetragen hast:"
    W ws, "S", SET_SHEET & ": Schuljahresstart, Stunden pro Woche, verfügbare Wochen, " & _
               "der komplette Schulwochen-Kalender mit allen Häkchen, die " & _
               "Ferientabelle und die vier Kopfangaben."
    W ws, "S", LB_SHEET & ": alle Zeilen bis zur Summe - Referenzcode, von, bis, " & _
               "Bezeichnung, Std. lt. Lehrplan und die Kompetenzerwartungen " & _
               "(Spalten A bis E und J)."
    W ws, "S", WP_SHEET & ": je Planzeile die Spalten E bis K und M, also " & _
               "Unterrichtswoche, Lehrplan-Code, Thema, Kompetenzen, Material, das " & _
               "Erledigt-Häkchen, die Art der Stunde und die Notizen."
    W ws, "N", "Nicht übernommen wird, was ohnehin berechnet wird: Referenz-Codes, " & _
               "Datum, KW, die Warnspalten, die Ferienzeilen und die Rechenspalten " & _
               "der " & LB_SHEET & ". Das baut die neue Mappe selbst auf - deshalb " & _
               "muss danach nichts mehr gedrückt werden."
    W ws, "N", "Ebenfalls nicht übernommen wird ein eigener Umbau des Blattaufbaus. " & _
               "Wer Spalten verschiebt oder Formeln umschreibt, verliert das beim " & _
               "nächsten Update."
    W ws, "N", "Der Import geht nur in eine Mappe, in der noch nichts geplant ist. " & _
               "Das ist Absicht: so kann ein Fehlklick nichts kosten."
    W ws, "-", ""

    W ws, "H3", "Wenn etwas klemmt"
    W ws, "S", "Die Schaltflächen reagieren nicht: Excel hat die Makros blockiert. " & _
               "Datei schließen, im Explorer rechts anklicken, Eigenschaften, unten " & _
               """Zulassen"" ankreuzen, wieder öffnen und ""Inhalt aktivieren"" wählen."
    W ws, "S", "Die Mappe muss als .xlsm gespeichert werden, sonst gehen beim " & _
               "Speichern alle Funktionen verloren."
    W ws, "S", "Farben oder Formeln sehen seltsam aus: """ & CTRL_SHEET & """, " & _
               """Einrichtung / Reparatur"" drücken. Das ist der Reparaturknopf, " & _
               "nicht der Update-Knopf - für eine neue Fassung ist ""Daten " & _
               "importieren"" zuständig."
    W ws, "S", "Eine Zelle lässt sich nicht bearbeiten: das ist Absicht - berechnete " & _
               "Spalten und Ferienzeilen sind gesperrt."
    W ws, "-", ""

    AnleitungStandSetzen
    ws.Range("B2").Select
End Sub


