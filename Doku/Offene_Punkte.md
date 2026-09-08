# Offene Punkte

Stand 06.09.2026.

> **Die offenen Befunde stehen jetzt im Issue-Tracker**, nicht mehr hier.
> Diese Datei behält nur, was sich schlecht als Issue führen lässt: bewusste
> Entscheidungen, den Migrationsverlauf und den Infrastrukturstand.
> Zwei Listen derselben Sache laufen sonst auseinander.

## Die Befunde der Gegenlese — jetzt als Issues

Ergebnis des ersten Einsatzes von Ebene 9. Alle fünf bestätigten Befunde wurden am
06.09.2026 unabhängig ein zweites Mal im Code nachgeprüft, bevor sie angelegt wurden;
die Zeilenangaben stimmen mit dem aktuellen Stand überein.

| # | Befund | Sicherheitsgrad |
|---|---|---|
| [#1](https://github.com/f-reiser/Stoffverteilungsplan/issues/1) | Leerzeilen-Prüfung hat denselben blinden Fleck wie der Fehler, den sie finden soll (Zeile 363) | sicher |
| [#2](https://github.com/f-reiser/Stoffverteilungsplan/issues/2) | `Chk ..., True` — Check prüft ein Literal (Zeile 255) | sicher |
| [#3](https://github.com/f-reiser/Stoffverteilungsplan/issues/3) | T3b prüft `wpBtnDown` statt `wpBtnUp` (Zeilen 547–548) | sicher |
| [#4](https://github.com/f-reiser/Stoffverteilungsplan/issues/4) | „Fehler 0" statt der echten Nummer (Zeile 1347) | sicher |
| [#5](https://github.com/f-reiser/Stoffverteilungsplan/issues/5) | Reparatur-Rat heilt Mutation 1 und 8 nicht | sicher |
| [#6](https://github.com/f-reiser/Stoffverteilungsplan/issues/6) | Sechs Checks ohne Mutation, Lücken in Abschnitt 10 und bei der PDF-Ausgabe | noch nicht nachgeprüft |
| [#7](https://github.com/f-reiser/Stoffverteilungsplan/issues/7) | `MappeIstLeer` prüft nur F/G/H/I/M — **entschieden: bleibt so**, ein Haken in J oder eine Kategorie in K allein ist keine Planung | kein Fehler |
| [#8](https://github.com/f-reiser/Stoffverteilungsplan/issues/8) | Vier veraltete Zahlen in `LIESMICH.txt` | sicher |

Neue Befunde entstehen über `.github/ISSUE_TEMPLATE/gegenlese-befund.yml`. Nicht
bestätigte Befunde werden **nicht** stillschweigend verworfen, sondern mit einer Zeile
Begründung geschlossen — sonst kommt derselbe Befund bei der nächsten Gegenlese wieder.

## Bewusst nicht gemacht

- `PrintTitleRows` wirkt nicht, weil die Titelzeile im Druckbereich liegt. Titelblock auf
  Seite 1 hat Vorrang. Der Nutzer weiß davon und kann es umdrehen lassen.
- `ANLEITUNG_STAND` muss bei jeder Textänderung am Anleitungsblatt hochgezählt werden,
  sonst bleibt ein vorhandenes Blatt mit altem Text stehen.

## Migration: erledigt und belegt

Der Nutzer zog alle Dateien auf den neuesten Stand (Module importieren, je Datei einmal
„Einrichtung / Reparatur" und „Wochenplan neu aufbauen"). Befund davor: VBA-Update-Blatt
in allen sechs Dateien, Leerzeilen in vier davon.

Gegenprobe am 05.09.2026 gelaufen:
`cd Makros && python3 pruefe_alles.py --alle-mappen` → **alle 8 Schritte grün**, über
Vorlage, Referenzmappe und alle fünf aktuellen Pläne. Beide Befunde sind weg.

**Korrektur vom 06.09.2026:** Dieser Lauf stand hier als „auf seinem Rechner". Das kann
nicht stimmen — `pruefe_module.py --selbsttest` schrieb nach `/tmp` und bricht unter
Windows ab. Der Lauf fand in der Linux-VM statt. Der Befund und seine Reparatur stehen
in `Doku/Testebenen.md`, Abschnitt „Was die Migration nach Claude Code aufgedeckt hat".

## Stand 06.09.2026 — Umzug nach Claude Code

Erster vollständiger Prüflauf auf dem **Windows-Rechner des Nutzers**: zunächst 7 von 8,
nach der Reparatur grün. Mit der Ausgabeschicht und der neuen Ebene A sind es jetzt
**11 Schritte, alle grün** — lokal wie in der CI.

Dabei repariert, beides in `pruefe_module.py`:
1. fester Pfad `/tmp` → `tempfile` (lief unter Windows überhaupt nicht),
2. die Arbeitskopie behält den Namen ihrer Quelldatei — vorher löste allein der
   Dateiname Regel 7 aus, wodurch zwei der sieben Mutationen nie etwas nachwiesen.

Neu: `Makros/ci_ausgabe.py` — Annotationen, Job Summary und Log-Gruppen für GitHub
Actions, 16 Prüfungen und 13 Mutationen, außerhalb der CI wirkungslos. Bewusst **kein**
`unittest`/`pytest`; Begründung in `Doku/Testebenen.md`, Abschnitt „Ebene W".

### Beinahe-Unfall beim ersten Push: Produktivdaten in `Vorlage/`

Beim Prüfsummenvergleich kurz vor dem Commit fiel auf, dass
`Vorlage/Referenzmappe.xlsm` **byteweise identisch** mit
`Stoffverteilungsplan_Mathematik_Gym_Bayern_Jgst10.xlsm` war — dem echten, produktiven
Plan. Ebenfalls betroffen: die *leere* Vorlage und `docProps/core.xml`. Gefunden wurden
Klarname, Schulname und der vollständige OneDrive-Pfad samt Windows-Benutzernamen im
versteckten Namen `wpPdfOrdner`.

`Offene_Punkte.md` hatte das vorhergesehen („Wenn dir die echten Plandaten im Repo nicht
passen, tausche `Referenzmappe.xlsm`…"), aber als Anregung, nicht als Prüfung — und wäre
damit beim Push untergegangen.

Beide Mappen sind jetzt erzeugt (`anonymisiere.py`) und werden bei jedem Push geprüft
(`pruefe_anonym.py`, Ebene A). Die Originale liegen unter
`Sicherung/Vorlage_mit_Echtdaten_2026-09-06/` (nicht im Repo).

**Damit ist auch die Template-Frage beantwortet:** `Vorlage/` läuft nicht mehr von der
Arbeitsdatei weg, weil es nicht mehr von Hand gepflegt, sondern neu erzeugt wird.

## Infrastruktur

Repo `f-reiser/Stoffverteilungsplan` existiert, privat. Von der Cowork-Umgebung aus war
GitHub nicht erreichbar (Gateway-Allowlist pro Sitzung; GitLab, Notion, Jira komplett
geblockt). Deshalb der Wechsel zu Claude Code. Versionierung, Changelog und Issue-Tracker
sollen über GitHub laufen.

**Die CI läuft.** `.github/workflows/pruefung.yml` ruft `Makros/pruefe_alles.py` —
dasselbe Skript, das lokal läuft. Abgedeckt sind die Ebenen 1, 2, 3, 6, 7, 8, A und W
samt ihrer Mutationstests, zusammen 61 Mutationen. Nicht abgedeckt sind Ebene 4/5
(brauchen echtes Excel) und Ebene 9 (braucht ein Sprachmodell).

## Der erste Push — erledigt am 06.09.2026

1. `git init` im Projektordner, Historie des bestehenden Repos übernommen (kein
   Force-Push, die beiden Ausgangs-Commits sind erhalten).
2. **`.gitignore` auf eine Whitelist umgestellt.** Erst wird alles ignoriert, dann
   gezielt zugelassen — Begründung in der Datei selbst und in `CLAUDE.md`.
   `modKonfig.bas` wurde vor dem Commit ausdrücklich gegengeprüft.
3. **`.gitattributes` ergänzt.** `core.autocrlf=input` hätte beim Commit CRLF nach LF
   umgeschrieben und dabei `Makros_verteilen.cmd` (Batch braucht CRLF) sowie die
   `.ps1` angefasst. Jetzt gilt `* -text`: Git fasst Zeilenenden nicht an.
4. **`Vorlage/` anonymisiert** — siehe oben.
5. **Acht Issues angelegt** (#1–#8), die fünf bestätigten Befunde vorher ein zweites
   Mal im Code nachgeprüft.
6. Erster CI-Lauf: **grün**, alle 11 Schritte.

Entschieden: `Bilder/` bleibt draußen. Der Code greift nicht mehr darauf zu — die Logos
liegen fest in der Mappe (`modKalender.bas:559`, `LIESMICH.txt:246`). Ebenso draußen
bleiben die Lambacher-Schweizer-`.docx` (Verlagsmaterial, 1,08 MB Binärdaten).

**Abgemachte Pflicht für die Zukunft:** Legt der Nutzer neue Verzeichnisse oder Dateien
an, muss vor dem Commit nachgefragt werden, ob sie ins Repository sollen. Die Whitelist
schützt vor Versehen in die eine Richtung — in die andere Richtung schützt nur die
Nachfrage.
