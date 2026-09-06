# Projektwissen — Aufbau der Mappe, Generator, Recherche

Fachliches Wissen, das nicht in `Makros/LIESMICH.txt` steht.

## Aufbau der Arbeitsmappe

Blätter nach dem Setup: `Steuerung`, `Wochenplan` (sheet2.xml), `Lernbereiche`
(sheet3.xml), `Einstellungen` (sheet4.xml), `Anleitung`.
(Das frühere sechste Blatt „VBA-Update"/„Update" ist am 05.09.2026 entfallen.)

**Wochenplan** — Titelblock Zeile 1–2, Überschriften Zeile 3, Daten ab Zeile 4:
A Warnkennzeichen · B Referenz-Code · C Datum · D KW · **E UW** · F Lehrplan-Code ·
G Thema · H Kompetenzen · I Material · J Erledigt (Kontrollkästchen, XF-basiert) ·
K Stunde (Dropdown aus `Einstellungen!M2:M6`) · L Hinweis · M Notizen ·
**N Kennzeichen** · O1 Sammelhinweis · P–S Warnspalten · U Lernbereich per FILTER ·
V Zeiger MATCH in die Wochentabelle. Zeilenhöhe 45,75.

**Ferienzeilen** stehen zwischen den Inhaltszeilen, Höhe 19,5: B:M verbunden, grauer
Hintergrund, fett/kursiv, **N = „F"** in der Hintergrundfarbe (unsichtbar) als Marker.
Ältere Mappen haben dort „FERIEN" — `MARK_TAG_ALT`, wird weiterhin akzeptiert.
Sie entstehen aus der Ferientabelle in `Einstellungen!I:K`, nicht aus den D-Flags.

**Lernbereiche** — Überschriften Zeile 3, Daten ab 4, direkt darunter die Summenzeile:
A Referenzcode · B von · C bis · D Lernbereich · E Std. lt. Lehrplan ·
**F–I Rechenspalten** · J Kompetenzerwartungen.

**Einstellungen** — B1 Schuljahresstart · B2 Stunden pro Woche · B3 verfügbare Wochen ·
B4 verfügbare Klassen-Wochen (Formel) · B5 fixiert (ja/nein) · A8–A11 Fach/Klasse/
Schule/Lehrkraft mit Werten in B · M2:M6 Kategorienliste (Stofferarbeitung,
Vorbereitung SA, SA/KA/kasL, Puffer, Verschiebbar) · Wochentabelle Zeilen 15–52 mit
D „Woche verfügbar", E „halbe Woche", F „halbe Klasse" (alle `t="b"`), G/H kumuliert,
I/J/K Ferientabelle.

**Bedingte Formatierung** auf B4:M{last}, alle Regeln nutzen `INDIRECT("X"&ROW())` und
sind daher zeilenunabhängig: 1 Erledigt, 2–6 Kategorie aus K, 7 halbe Klasse,
8 halbe Woche. Sieben davon sind x14-Regeln.

## Die Mappe bearbeiten

**NIE mit openpyxl speichern** — es wirft die x14-Erweiterungen (bedingte Formatierung,
Dropdown für K) ersatzlos weg und beschädigt `xl/vbaProject.bin`. Lesen ist unproblematisch.

Zum Schreiben: Zip entpacken, `xl/worksheets/sheet*.xml` per String-Ersetzung anpassen,
neu packen (`[Content_Types].xml` zuerst). Neue Textzellen als `t="inlineStr"`, dann
bleibt `sharedStrings.xml` unangetastet.

Ändert sich die Zeilenzahl, müssen ALLE Bereichsangaben mit: `<dimension>`,
`conditionalFormatting sqref`, `dataValidations sqref`,
`x14:conditionalFormatting/xm:sqref`, `x14:dataValidation/xm:sqref` und der
COUNTIF-Bereich in O1. Sonst hören Farben und Dropdowns unter der alten letzten Zeile auf.

Nach jeder Änderung `xl/calcChain.xml` löschen (plus Eintrag in `[Content_Types].xml`
und `xl/_rels/workbook.xml.rels`) und in `xl/workbook.xml` `fullCalcOnLoad="1"` setzen.

Achtung bei `re.sub`: Backslash-Sequenzen im Ersetzungsstring werden interpretiert —
Lambda als Ersetzung verwenden. **Style-Indices sind PRO DATEI unterschiedlich** — nie
Style-Nummern aus einer Notiz übernehmen, immer aus der konkreten Quelldatei auslesen.

Prüfliste nach jedem Eingriff: ZIP-Integrität, XML-Wohlgeformtheit aller Teile,
`vbaProject` vorhanden, `featurePropertyBag` vorhanden, `calcChain` weg,
`fullCalcOnLoad` gesetzt, 8 CF-Regeln (per `<cfRule ` bzw. `<x14:cfRule ` zählen, NICHT
per Substring — der matcht Open- und Close-Tag doppelt), Zeilen lückenlos, kein
`#REF!`/`_xludf`.

## LibreOffice-Validierung

`soffice --headless --convert-to ods x.xlsm`, dann `--convert-to xlsx ...ods`
(zweistufig — einstufig bleiben Cross-Sheet-Formeln unkalkuliert). Bekannte, HARMLOSE
Abweichungen: `_xlfn._xlws.FILTER` und `_xlfn.SEQUENCE` liefern `#NAME?`, Spalte Q zeigt
`#NAME?` (REGEXTEST unbekannt), Spalte U bleibt leer, „TT.MM" bleibt Literaltext,
`COUNTIFS(...;TRUE())` zählt alles mit, einzelne `SUM()` bleiben 0. Alles auch in einer
unveränderten, produktiv genutzten Datei des Nutzers reproduzierbar — also nicht
generatorspezifisch. Verlässlich ist der Soll/Ist-Vergleich der reinen Wertezellen
(openpyxl, `data_only=False`).

## Kalenderdaten

Der Template-Kalender weicht an zwei Stellen von der Realität ab, bestätigt durch
Abgleich mit allen realen Fachdateien: Woche 1 (14.–20.09.) ist verfügbar, aber als
halbe Woche (D=TRUE, E=TRUE); Woche 17 (25.–31.01.) ist NICHT verfügbar (D=FALSE).
Schulweit, nicht fachspezifisch.

## LehrplanPLUS auslesen

`WebFetch` liefert bei `lehrplanplus.bayern.de` nur den Anfang — bei Mathematik FOS 11
brach es mitten in Lernbereich 2 ab, und das war aus dem Ergebnis nicht erkennbar.
Der `/schulart/...`-Pfad ist per robots.txt gesperrt, nur `/fachlehrplan/...` geht.

Vorgehen: Seite im Browser-Pane öffnen, dann
`document.querySelector('a.paragraph_toggle_all').click()`, ~2 s warten, `get_page_text`.
URL-Muster: `https://www.lehrplanplus.bayern.de/fachlehrplan/<schulart>/<jgst>/<fach>/<ausprägung>`
— z. B. `.../fos/11/mathematik/abu-g-s-w-gh-iw_gueltig_ab_26_27` oder
`.../gymnasium/9/mathematik`.
Gegenprobe: Die Übersicht ohne Aufklappen listet alle Lernbereiche mit Zeitrichtwerten.

## Lambacher Schweizer 9 (Bayern, 733091) — Seitenzuordnung

Kapitelgrenzen: I 4–33 · II 34–73 · III 74–101 · IV 102–117 · V 118–141 · VI 142–161 ·
VII 162–183 · VIII 184–215.

Die Unterkapitelgrenzen wurden aus dem Lösungs-PDF gewonnen (`pdftotext -layout`;
Unterkapitel-Überschriften sind durch ein `\x07` zwischen Nummer und Titel markiert).
Anfangsseiten belastbar, Endseiten als „Anfang des nächsten minus 1" — an Übergängen
kann eine Seite danebenliegen. Vom Nutzer noch nicht gegengeprüft.
Die vollständige Tabelle steht im Projektgedächtnis der Cowork-Sitzung; bei Bedarf neu
herleiten.

## Erzeugte Dateien

Template plus fünf Pläne: Mathematik Gym 9/10/11, Physik Gym 9, Mathematik FOS 11.
Alte Fassungen liegen als `... - v1.xlsm` bzw. `... - old.xlsm` daneben — das sind die
Archive des Nutzers, nur lesen.
