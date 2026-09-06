# Offene Punkte

Stand 05.09.2026, Ende des Tages.

## Vom unabhängigen Agenten gemeldet, von mir im Code bestätigt, NOCH NICHT repariert

Der Nutzer wollte das ausdrücklich als Issues behandeln, nicht als stille Reparaturen.
Alle fünf habe ich selbst im Code nachgeprüft — Sicherheitsgrad **sicher**.

**1. Die Leerzeilen-Prüfung hat immer noch denselben blinden Fleck.**
`modSelbsttest.T1b_Zustand`, Zeile ~363: `If Not NurLeer(ws, r) Then Exit Do` ist
unterhalb von `PlanLastRow` immer wahr, weil `PlanLastRow` dieselben Spalten scannt wie
`NurLeer`. Entschieden wird allein über die Zeilenhöhe — eine Formatheuristik, während
`LeereEndzeilenEntfernen` `UsedRange` benutzt. Test und Code messen verschiedene Dinge.
Die unabhängige Prüfung wäre `UnterkanteBlatt(ws) <= PlanLastRow(ws)`.

**2. Ein Check, der ein Literal prüft.**
`modSelbsttest`, Zeile ~255: `Chk "Fixierung ist ...", True`. Kann nicht fehlschlagen.

**3. T3b prüft den falschen Knopf.**
Der eigene Kommentar sagt „ReadBlock liest den Bereich von `wpBtnUp`" — geprüft wird
`FormBereich(ws, "wpBtnDown")`, den niemand liest. Nähme man die `AlternativeText`-Zeile
aus `HideOne` heraus, bliebe der Selbsttest grün.

**4. „Fehler 0" statt der echten Nummer.**
Im `Fail:`-Zweig von `Selbsttest_Pruefen` steht `Err.Number` NACH `BerichtSchreiben`,
das `Err.Clear` aufruft. Im anderen Fail-Zweig ist die Falle kommentiert und vermieden,
hier nicht.

**5. Der Reparatur-Rat nach einem Abbruch ist falsch.**
Die Meldung sagt „Einrichtung / Reparatur drücken". Das heilt Mutation 1 und 8 nicht:
`Setup_Stoffverteilungsplan` ruft `LeereEndzeilenEntfernen` gar nicht auf und schreibt
Spalte E nicht neu.

## Weiteres aus der Gegenlese, noch nicht nachgeprüft

- Sechs Checks in T0/T1b haben keine zugehörige Mutation — für die ist unbewiesen, dass
  sie anschlagen: „Application.Run erreicht ein privates Modul", „Gegenprobe modStart",
  „Zeilen-Schaltflächen haben ein Makro", „Summenzeile gefunden", „Summenzeile steht
  direkt unter dem letzten Lernbereich", „Blatt Anleitung vorhanden".
- Mutation 2 leert nur Spalte F der ERSTEN Lernbereichszeile; G/H/I und die übrigen
  Zeilen sind unbewiesen. Mutation 5 legt nur „Update" an, nie „VBA-Update".
- Abschnitt 10 prüft die Übernahme der **Einstellungen** gar nicht: `MappeLeeren` räumt
  nur Wochenplan und Lernbereiche leer, Kalender/Ferientabelle/Kopfangaben bleiben stehen.
  Der Check „Schulwochen-Kalender uebernommen" vergleicht deshalb zwei Seiten derselben
  unveränderten Tabelle und ist immer grün.
- Kopf-/Fußzeilen im PDF haben keinen einzigen Check, obwohl der `&Z`-Vorfall die
  ausführlichste Warnung in der LIESMICH ist. Ebenso ungeprüft: `FitToPagesWide`,
  `PaperSize`, die Seiteneinrichtung des Blattes Lernbereiche, das Aus- und
  Wiedereinblenden der übrigen Blätter beim Export.
- `MappeIstLeer` prüft nur F, G, H, I, M — eine Mappe, in der jemand nur Spalte K oder J
  gefüllt hat, gilt als leer und wird ohne Rückfrage überschrieben. **Auslegungsfrage,
  der Nutzer muss entscheiden.**
- Vier Zahlen in `LIESMICH.txt` sind veraltet: vier Alt+F8-Einträge (es sind fünf),
  „rund 80 Prüfungen in acht Abschnitten" (129 in dreizehn), „vier direkte MsgBox in
  modSelbsttest" (sieben), die Schaltflächenliste des Blattes Steuerung nennt sechs
  statt sieben Knöpfe.

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
nach der Reparatur **alle 9 Schritte grün** (der neunte ist die Ausgabeschicht).

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

**Die CI ist fertig und liegt bei.** `.github/workflows/pruefung.yml` ruft
`Makros/pruefe_alles.py` — dasselbe Skript, das lokal läuft. Abgedeckt sind die Ebenen
1, 2, 3, 6, 7 und 8 samt ihrer Mutationstests, zusammen 40 Mutationen. Nicht abgedeckt
sind Ebene 4/5 (brauchen echtes Excel) und Ebene 9 (braucht ein Sprachmodell).

## Was beim ersten Push zu tun ist

1. **`git init` im Ordner `Stoffverteilungsplan/`** — nicht in `Makros/`. Die Doku, die
   CI-Konfiguration und die Referenzmappe liegen außerhalb von `Makros/`.
2. **`.gitignore` liegt bei.** Sie schließt `modKonfig.bas` (Kennwort), alle `.xlsm`
   außer `Vorlage/`, `Diagnose/`, `Sicherung/` und die Excel-Sperrdateien aus.
   *Vor dem ersten Commit `git status` lesen und prüfen, dass `modKonfig.bas` wirklich
   nicht dabei ist.*
3. **`Vorlage/` ist schon angelegt** und enthält beides: die leere
   `Stoffverteilungsplan_Template.xlsm` und `Referenzmappe.xlsm` (eingefrorene Kopie von
   Mathematik Gym 10). Die zweite treibt die Mutationstests — in einer leeren Vorlage
   greift keine Mutation. Beide nur bewusst austauschen. Wenn dir die echten Plandaten
   im Repo nicht passen, tausche `Referenzmappe.xlsm` gegen eine ausgedünnte Mappe;
   nötig sind nur ein paar gefüllte Planzeilen.
4. **Einmal lokal `cd Makros && python3 pruefe_alles.py`** und erst pushen, wenn das
   grün ist. Sonst ist der erste CI-Lauf rot, und niemand weiß, ob an der Pipeline oder
   an der Sache.
5. **Aus den fünf bestätigten Befunden oben fünf Issues machen** — Vorlage
   `.github/ISSUE_TEMPLATE/gegenlese-befund.yml`. Sie sind bewusst nicht repariert
   worden; der Nutzer wollte sie als Issues.

Offen und vom Nutzer zu entscheiden: ob `Bilder/` (Schullogos) mit ins Repo soll.
