# Bestätigte Fallstricke

Alles hier ist mindestens einmal aufgeschlagen — meist beim Nutzer, was jedes Mal
vermeidbar gewesen wäre. Ausführlicher, mit Codebeispielen, in `Makros/LIESMICH.txt`.

## Excel-Objektmodell

**Bedingte Formatierung wird per VBA nur gezählt, plus EIN Sonderfall.** Erlaubt sind
`FormatConditions.Count` und, seit PR #59 ("Weg 2"), `AppliesTo.Areas.Count` einer
EINZELNEN Regel über einen festen, literalen Index — das erkennt eine Fläche, die in
mehrere Teilbereiche zerfallen ist, obwohl die Regelzahl je Zeile noch stimmt. Kein
`FormatConditions.Add`, kein `ModifyAppliesToRange`, keine Enumeration (auch nicht über
einen Schleifenindex), kein `Formula1`. Sieben der acht Regeln sind x14-Erweiterungsregeln;
der Zugriff auf die klassische `FormatConditions`-Auflistung hat Excel am 31.08.2026 hart
abstürzen lassen. Das Zählen ist seit dem 13.09.2026 frei (Issue #31): `modPruefung`
erkennt daran, ob der Farbbereich noch alle Planzeilen erreicht.
Neue Zeilen stattdessen über „kopierte Zellen einfügen" erzeugen
(`Rows(r).Copy`, dann `Rows(r+1).Insert Shift:=xlDown`) — Excel überträgt Formate,
Kontrollkästchen-XF, Zeilenhöhe, Gültigkeitsliste UND bedingte Formatierung selbst mit.
Folgeregel: Bereiche mit CF oder Gültigkeit werden AUSGEBLENDET statt verschoben.

**Verbundene Zellen.** Nie eine Eigenschaft auf einen Bereich setzen, der eine verbundene
Zelle nur TEILWEISE überdeckt — Ferienzeilen sind B:M verbunden, `Range("E2:K37").Locked`
gibt Laufzeitfehler 1004. `Rows(r).MergeCells` liefert bei teilweiser Verbindung Null →
Fehler 94 beim Zuweisen an Boolean. Und `Range.Characters(...).Font` wirkt auf
verbundenen Zellen NICHT.

**Grafiken.** `src.Copy` + `ws.Paste` RASTERT ein Bild in Anzeigegröße neu (714×80 wurde
zu 66×55). Nie kopieren — einmal einbetten, nur ein-/ausblenden und skalieren. Beim
Skalieren NUR die Höhe setzen. Bilder hängen nicht an Zellen; Ausblenden nur über
`Shape.Visible = msoFalse`. `Placement`: `xlMoveAndSize` zieht eine Form auf die volle
Zeilenhöhe, sobald die Zeile wächst — für Schaltflächen in hohen Zeilen `xlMove`.

**PageSetup / PDF.** `Workbook.ExportAsFixedFormat` exportiert IMMER die ganze Mappe;
für eine Teilmenge die übrigen Blätter kurz `xlSheetHidden` setzen (auch im Fehlerzweig
zurücksetzen). `SelectedSheets.ExportAsFixedFormat` ist ein Compile-Fehler. Unter
`PrintCommunication = False` kommen Kopf-/Fußzeilen NICHT an. `&Z` ist der volle PFAD der
Datei — stand einmal als SharePoint-Adresse über jeder PDF-Seite.
`PrintTitleRows` wirkt NICHT, wenn die Titelzeile INNERHALB des Druckbereichs liegt; hier
beginnt der Druckbereich bei Zeile 1 (Titelblock), die Überschrift wird also nicht
wiederholt. Beides zugleich geht nicht, Titelblock hat Vorrang. Dem Nutzer bekannt.

**`Dir$`-Reentranz.** Genau EIN Suchzustand für die ganze Anwendung; ein weiteres `Dir$`
in einer laufenden Schleife setzt diese zurück.

## Die drei lautlosen Fallen

**1. `UsedRange` und `End(xlUp)` beantworten verschiedene Fragen.** `End(xlUp)` findet, wo
der INHALT endet; `UsedRange`, wo das BLATT endet — inklusive leerer, nur formatierter
Zeilen. `LeereEndzeilenEntfernen` begann bei `PlanLastRow` (End(xlUp)) und suchte nach
oben; 35 formatierte Reservezeilen lagen DARUNTER und waren unsichtbar.
Verwandt: In einer LEEREN Mappe liefert alles, was auf `End(xlUp)` beruht, immer die
erste Datenzeile. Eine Abbruchbedingung darauf zu bauen hat den Wochenplan-Import nach
genau einer Zeile beendet.

**2. `Worksheets(...).Delete` läuft gegen `ThisWorkbook.Protect Structure:=True` ins
Leere** — unter `On Error Resume Next` völlig lautlos. Alles, was die MAPPENSTRUKTUR
ändert, braucht `ThisWorkbook.Unprotect SCHUTZ_PW` und danach wieder `Protect`. Und:
nach dem Löschen NACHPRÜFEN und das Ergebnis zurückgeben.

**3. `ThisWorkbook.Path` ist nicht immer ein Pfad.** Bei eingeschaltetem AutoSpeichern
meldet Excel die SharePoint-ADRESSE (`https://...`). `Open ... For Output` scheitert dort,
und `C:\Temp` gibt es auf einem normalen Windows nicht. Rückfallkette: Mappenordner,
`%TEMP%`, Dokumente, `C:\Temp`.

## VBA-Syntax, die einen Compile-Fehler gibt

- **Modulweite Deklarationen** (`Private Const`, `Private m...`) gehören AUSSCHLIESSLICH
  in den Deklarationsteil ganz oben. Zwischen zwei Prozeduren: „Nach End Sub … können nur
  Kommentare stehen", gefolgt von „Variable nicht definiert". Innerhalb einer Prozedur
  ist `Dim` natürlich erlaubt.
- `ActiveWindow.SelectedSheets.ExportAsFixedFormat` — die Methode gibt es nur auf
  Workbook, Worksheet und Range.
- `xlNoRestriction` statt `xlNoRestrictions` — jede `xl*`/`mso*`-Konstante gegen das
  Objektmodell prüfen, nie per Präfix durchwinken.
- Farben nie als vorausgerechnete Long-Konstante — `Const` kann `RGB()` nicht aufrufen,
  und die Handrechnung war ZWEIMAL falsch. `Public Function FARBE_KOPF() As Long`.

## Sonstiges

`FastOn`/`FastOff` sind NICHT schachtelbar: `FastOn` steigt bei `mBusy` aus, das erste
`FastOff` schaltet den Blattschutz wieder ein. Wer aus einem laufenden Block heraus eine
Routine mit eigenem Paar ruft, bekommt beim nächsten Schreibzugriff Fehler 1004.
Außerdem: vor `FastOn` NICHT von Hand `ScreenUpdating = False` setzen — `FastOn` merkt
sich genau diesen Wert und `FastOff` stellt ihn wieder her.

**Alt+F8 zeigt die Makros ALLER geöffneten Mappen**, jeweils als `'Datei.xlsm'!Makro`.
Der Nutzer hat das einmal für doppelte Module gehalten. Gegenprobe an der Datei:
`vbaProject.bin` entpacken und die Modulnamen im OLE-Verzeichnis zählen (als UTF-16
im Binary suchbar).

**Meldungen:** Windows spielt den Systemklang wegen des SYMBOLS. `MsgBox` ohne
`vbInformation`/`vbExclamation`/`vbQuestion`/`vbCritical` bleibt stumm — das macht
`modWochenplan.OhneTon`.

**Dateizugriff:** Die xlsm ist gesperrt, solange sie in Excel offen ist. Löschen im
gemounteten Ordner ist nicht erlaubt, Überschreiben schon. Eine Sperrdatei
`~$Name.xlsm` verrät, ob die Datei offen ist.

## Bestätigt funktionsfähig

`Range.Formula` mit US-Namen und `_xlfn`-Präfixen, `.Formula2`, `Validation.Add` mit
Fremdblattbezug, `Application.GetSaveAsFilename` (liefert bei Abbruch `False` — mit
`VarType` prüfen), benannte Bereiche als Brücke zwischen Blättern (`wpPdfFormat`,
`wpAnleitungStand`), `ChrW$(9888)` als Warnzeichen im Schaltflächentext,
`Option Private Module` ohne Wirkung auf `OnAction`.
