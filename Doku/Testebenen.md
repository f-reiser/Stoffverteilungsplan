# Die neun Prüfebenen

Aufgebaut am 05.09.2026, nachdem der Nutzer gefragt hatte, ob die Tests überhaupt
validiert sind. Die Antwort war damals ehrlicherweise „nur teilweise" — mit Belegen.
Am selben Tag in Code gegossen, nachdem er darauf bestand: die Ebenen 6, 7 und 8 waren
bis dahin **Gewohnheiten** — Skripte, die in einer Sitzung entstanden und mit ihr
verschwanden. Eine Gewohnheit kann niemand erben.

## Warum es diese Ebenen gibt

Drei Belege dafür, dass ungeprüfte Tests wertlos sind, alle aus diesem Projekt:

- Der Check „Keine leeren Zeilen am Ende des Plans" meldete **ok**, während 35 leere
  Zeilen unter der Tabelle standen. Er begann bei `PlanLastRow` und suchte nach oben —
  die Zeilen lagen darunter. **Der Check hatte denselben blinden Fleck wie der Fehler,
  den er finden sollte.**
- Die Python-Import-Simulation meldete gleichzeitig „0 Abweichungen", weil sie nur den
  Bereich betrachtete, in den geschrieben wurde.
- Der Check „Summenformel zeigt auf die Datenzeilen" war grün, obwohl nichts sie neu
  gebaut hatte — die alte Formel hatte zufällig überlebt. Falsches Grün.

**Regel daraus:** Wenn der Code eine Größe mit einer Funktion bestimmt, darf der Test sie
nicht mit derselben Funktion nachprüfen. Sonst bestätigt der Test nur, dass der Code zu
sich selbst passt. Vor jedem Test fragen: *woher weiß der Test das, und woher weiß der
Code es?*

---

## Ein Aufruf für alles

```
cd Makros
python3 pruefe_alles.py
```

Das ist derselbe Aufruf, den `.github/workflows/pruefung.yml` bei jedem Push macht.
Absichtlich dasselbe Skript: so können Arbeitsplatz und Pipeline nicht auseinanderlaufen,
und eine neue Prüfung landet mit einem einzigen Eintrag an beiden Stellen.

Rückgabewert 0 = alles grün. Am Ende steht immer eine Zusammenfassung Zeile für Zeile,
und darunter, was **nicht** geprüft werden konnte.

`--ohne-mappen` überspringt alles, was eine `.xlsm` braucht.

---

## Die Ebenen

| # | Was | Aufruf | Umfang | Läuft in der CI |
|---|---|---|---|---|
| 1 | statische Prüfung der Module | `vbacheck.py mod*.bas` | 9 Regelgruppen | ja |
| 1M | Mutationstest über Ebene 1 | `vbacheck.py --selbsttest mod*.bas` | 9 Mutationen | ja |
| 2 | Abnahme fertiger `.xlsm` von außen | `pruefe_datei.py ../*.xlsm` | 7 Inhalts- + 7 Strukturregeln | ja¹ |
| 3 | Mutationstest über Ebene 2 | `pruefe_datei.py --selbsttest` | 10 Mutationen | ja¹ |
| 4 | `modSelbsttest.Selbsttest` in echtem Excel | Schaltfläche in der Mappe | 139 Chk-Aufrufe, 13 Abschnitte | **nein** — braucht Excel |
| 5 | Mutationstest über Ebene 4 | Schaltfläche „Selbsttest prüfen" | 10 Mutationen | **nein** — braucht Excel |
| 6 | Formelkonsistenz aller Planzeilen | `pruefe_formeln.py ../*.xlsm` | 12 Spalten, 4 Regeln | ja¹ |
| 6M | Mutationstest über Ebene 6 | `pruefe_formeln.py --selbsttest` | 4 Mutationen | ja¹ |
| 7 | Auslieferungsprüfung der Quelldateien | `pruefe_module.py mod*.bas *.txt` | 8 Kriterien je Datei | ja |
| 7M | Mutationstest über Ebene 7 | `pruefe_module.py --selbsttest` | 7 Mutationen | ja |
| 8 | Strukturprüfung der Datei (Zip/XML) | in `pruefe_datei.py` enthalten | 7 Regeln | ja¹ |
| 9 | Fremde Gegenlese durch unabhängige Agenten | `Doku/Ebene9_Gegenlese.md` | Ablauf, kein Test | **nein** — braucht ein Modell |
| W | Ausgabeschicht für die CI | `ci_ausgabe.py --selbsttest` | 16 Prüfungen, 13 Mutationen | ja |
| A | keine Produktivdaten in `Vorlage/` | `pruefe_anonym.py ../Vorlage/*.xlsm` | allgemeine + standortabhängige Spuren | teilweise¹ |
| AM | Mutationstest über Ebene A | `pruefe_anonym.py --selbsttest` | je Spur eine Mutation | teilweise¹ |

¹ nur, wenn eine `.xlsm` im Repo liegt — siehe „Die Referenzmappe".
² Ebene A läuft in der CI, prüft dort aber nur die **allgemeinen** Spuren: Klarnamen und
Schulnamen stehen in `anonym_muster.local.json`, das `.gitignore` ausschließt und das in
der CI deshalb nie existiert. Vor dem Push ist der lokale Lauf die eigentliche Prüfung —
fehlt die Datei dort, bricht sie mit einem Fehler ab.

Dazu die Abnahme des Nutzers in echtem Excel. Die hat bisher am meisten gefunden und ist
durch nichts hiervon zu ersetzen.

**Jede Ebene hat einen Mutationstest, oder sie zählt nicht.** Eine Prüfung ohne Nachweis,
dass sie überhaupt anschlagen kann, ist eine Behauptung. Dieses Projekt hat zwei Prüfungen
gehabt, die strukturell nie anschlagen konnten — beide fielen erst durch ihren eigenen
Mutationstest auf.

## Ebene A — keine Produktivdaten im Repository

`Vorlage/Referenzmappe.xlsm` war bis 06.09.2026 eine **byteweise Kopie** des echten
Mathematik-Gym-10-Plans. Aufgefallen ist das erst beim Vergleich der Prüfsummen, wenige
Minuten vor dem ersten Push. Enthalten waren: Themen, Termine, Schulaufgaben, eigene
Notizen, der Klarname der Lehrkraft, der Schulname — und im versteckten Namen
`wpPdfOrdner` der vollständige OneDrive-Pfad samt Windows-Benutzernamen. Dieselben
Klarnamen steckten auch in der *leeren* Vorlage und in `docProps/core.xml`
(`lastModifiedBy`).

Beide Mappen werden deshalb aus einer produktiven Datei **erzeugt**
(`anonymisiere.py`, reine XML-Chirurgie im entpackten Zip — openpyxl darf sie nie
speichern). Damit können sie auch nicht mehr unbemerkt von der Arbeitsdatei abweichen.

`pruefe_anonym.py` durchsucht **jeden Teil** der Mappe, auch `xl/vbaProject.bin`, in
cp1252 wie in UTF-16LE. Die frühere Ausnahme für den VBA-Strom — „dort steht ohnehin nur
der Quelltext, der als `.bas` im Repository liegt" — hat zwei Lecks durchgelassen:
`modKonfig` wird nie als `.bas` ausgeliefert, steckt aber einkompiliert in jeder Mappe,
und eine nicht neu gebaute Mappe enthält weiter die **alte** Fassung eines geänderten
Moduls. So lagen Blattschutz-Kennwort und Schulnamen im Repository, während die Prüfung
„SAUBER" meldete.

Absichtlich **nicht** gemeldet werden Lernbereichsnamen und Lehrplan-Codes — die stehen so
im LehrplanPLUS und sind öffentlich. Gemeldet wird, was den Nutzer, seine Schule oder seine
konkrete Unterrichtsplanung erkennbar macht.

Die Prüfung bekommt bewusst eine **feste Dateiliste** (`Vorlage/`), auch bei
`--alle-mappen`. Die eigenen Pläne des Nutzers enthalten selbstverständlich echte Daten;
sie mitzuprüfen wäre eine dauerhaft rote Meldung ohne Aussage — und die liest bald
niemand mehr.

**Warum als Prüfung und nicht als Vorsatz:** Was einmal in der Git-Historie steht,
bekommt man nicht mehr sauber heraus. Eine Zusage, beim nächsten Mal daran zu denken,
läuft nicht bei jedem Push mit.

## Die Referenzmappe

Die fertigen `.xlsm` sind Binärdateien: nicht vergleichbar, nicht zusammenführbar, und
jedes Speichern ändert die halbe Datei. Sie gehören nicht in die Historie, und
`.gitignore` schließt sie aus.

Ausgenommen ist `Vorlage/`. Dort liegen **zwei** Dateien, und der Unterschied ist wichtig:

| Datei | wofür |
|---|---|
| `Stoffverteilungsplan_Template.xlsm` | die leere Vorlage. Abnahme (Ebene 2/8) und Formelkonsistenz (Ebene 6) prüfen sie mit. |
| `Referenzmappe.xlsm` | eine eingefrorene Kopie eines **gefüllten** Plans. Nur sie treibt die Mutationstests der Ebenen 3 und 6. |

**Warum es die zweite braucht — und das ist beim ersten Lauf schiefgegangen:** Die
Mutationstests bauen einen Fehler in eine Mappe ein und verlangen, dass er auffällt. In
einer leeren Vorlage gibt es nichts zu verbiegen; beide Prüfer erkennen die leere Mappe
und überspringen die Inhaltsregeln. Ergebnis beim ersten Lauf gegen die echte Vorlage:
**6 von 10** und **0 von 4**. Die Tests haben sich richtig verhalten — falsch war die
Referenzmappe. `pruefe_alles.py` sagt das jetzt im Klartext, statt vier rätselhafte
FEHLT-Zeilen zu drucken.

Beide Dateien werden nur bewusst ausgetauscht. Die eigenen Pläne prüft man mit
`--alle-mappen`; die Archive (` - v1`, ` -v2`, ` - old`, ` - Test`) bleiben dabei außen
vor, weil sie absichtlich auf altem Stand sind und eine dauerhaft rote Prüfung bald
niemand mehr liest.

## Das Kennwort

`Makros/modKonfig.bas` enthält das Blattschutz-Kennwort und ist in `.gitignore`. Ohne ein
Modul dieses Namens kennt aber kein anderes Modul die Konstante `SCHUTZ_PW`, und
`vbacheck.py` meldet an 14 Stellen zu Recht „nicht deklariert".

Deshalb liegt daneben `Makros/modKonfig.bas.vorlage` mit demselben Modulnamen und einem
Platzhalter-Kennwort. `pruefe_alles.py` kopiert die Module in ein Arbeitsverzeichnis und
setzt die Vorlage ein, wenn das echte Modul fehlt. Ausgeliefert wird die Vorlage nie —
die Endung `.vorlage` hält sie aus jedem `mod*.bas`-Glob heraus.

---

## Ebene 1 — vbacheck.py

Regeln: (1) vor der ersten Prozedur nur Deklarationen, (1b) und danach keine modulweite
Deklaration mehr, (2) Blockbalance, Sprungmarken je Prozedur gesammelt, (3) **Option
Explicit JE PROZEDUR** — nicht global, (4) kein FormatConditions-Zugriff, (5)
Office-Konstanten gegen eine explizite Liste, (6) `modXxx.Name` muss existieren UND dort
Public sein, (7) Liste verbotener Aufrufe, (8) Merge-Sicherheit bei Bereichen quer durch
B..M.

Der Mutationstest (`--selbsttest`) baut für jede dieser Gruppen einen Fehler in
`modWochenplan.bas` ein und verlangt, dass genau die zuständige Meldung kommt. Er weigert
sich zu laufen, wenn die Ausgangslage nicht sauber ist — sonst wäre nicht unterscheidbar,
ob eine Meldung von der Mutation stammt.

Jede neue Regel bekommt eine Mutation. Ohne Mutation keine Regel.

## Ebene 6 — pruefe_formeln.py

Naheliegend wäre, die `Frm*`-Funktionen aus `modWochenplan` in Python nachzubauen und
zeichengenau zu vergleichen. Als Dauertest taugt das nicht: es wäre eine zweite Umsetzung
derselben Idee, die mit jeder VBA-Änderung nachgezogen werden müsste — und wer beides aus
derselben Vorstellung schreibt, bekommt zweimal denselben Fehler. Genau der blinde Fleck,
gegen den diese Ebenen gebaut sind.

Die Prüfung stellt deshalb eine andere Frage, die ohne Kenntnis der VBA-Logik auskommt:
**„Sehen alle Planzeilen gleich aus?"** Eine Zeile, die aus der Reihe fällt, ist immer ein
Fehler — egal welche Formel richtig wäre. Dafür wird jede Formel normiert (Zeilennummer →
`@Z`, erste Planzeile → `@E`, fremde Blattbezüge → `@FREMD`) und mit der Mehrheit
verglichen.

Geprüft: Wochenplan `A C D L P Q R V`, Lernbereiche `F G H I`. Bewusst ausgenommen sind
`B`, `S` und `U` — dort steht je Zeile etwas anderes; die Begründung steht im Skript.
Dazu: Ferienzeilen dürfen in `C` keine Formel haben, und der Querbezug
`Lernbereiche!$A$n` wird aus dem rohen XML gelesen, weil openpyxl Spalte U nicht liefert.

Zwei Fehler im Prüfer selbst fielen dabei auf, beide **weil zuerst gegen bekannt gute
Dateien gefahren wurde**: eine nichtdeterministische Normierung (zwei Ersetzungspaare in
einem `set`) und eine zu gierige Zahlenersetzung, die `Einstellungen!$B$15`, `CHAR(10)`
und `>=9` zerlegte. Merksatz: **Erst grün gegen Bekanntes, dann rot gegen Mutationen.**

## Ebene 7 — pruefe_module.py

Kodierung und Auslieferbarkeit jeder `.bas`/`.txt`: als cp1252 lesbar, keine
UTF-8-Bytefolgen, kein `U+FFFD`, keine Doppelkodierung, keine doppelten Backslashes in
Zeichenketten, einheitliche Zeilenenden, `Attribute VB_Name` passend zum Dateinamen,
`Option Explicit`, Abschluss mit Zeilenumbruch.

Das ist keine Theorie: das Edit-Werkzeug hat `modAnleitung.bas` einmal als UTF-8
zurückgeschrieben und dabei jeden Umlaut durch `U+FFFD` ersetzt; gerettet wurde die Datei
nur, weil die Vorgängerfassung noch beim Nutzer lag.

**Was der Mutationstest hier über sich selbst gefunden hat:** Die `U+FFFD`-Prüfung konnte
nie anschlagen. Gesucht wurde das *Zeichen* U+FFFD im cp1252-gelesenen Text — dort kann es
das nicht geben, weil U+FFFD in der Datei als die drei Bytes `EF BF BD` steht und cp1252
die als `ï¿½` liest. Die Prüfung war seit ihrer Entstehung wirkungslos. Jetzt wird die
**Bytefolge** gesucht.

## Ebene 8 — Strukturprüfung (in pruefe_datei.py)

Zip-Integrität, Wohlgeformtheit jedes XML-Teils, `vbaProject.bin` vorhanden,
`featurePropertyBag` vorhanden, `calcMode`/`fullCalcOnLoad` stimmig, 1 klassische + 7
x14-Regeln der bedingten Formatierung, kein `#REF!` und kein `_xludf`.

Zwei Regeln sind hier wieder **gestrichen** worden, weil sie an echten Dateien
Fehlalarme gaben:
- `calcChain.xml` muss weg — das ist eine **Bau**regel nach XML-Chirurgie, keine
  **Abnahme**regel. Excel legt die Datei beim Speichern völlig zu Recht neu an.
- Die CF-Zählung zog `"<x14:cfRule "` von `"<cfRule "` ab. Das sind verschiedene
  Zeichenketten; die Subtraktion war schlicht falsch.

Merksatz: **Eine Regel, die an einer bekannt guten Datei anschlägt, ist keine Regel.**

## Ebene 5 — der Mutationstest in Excel, und was er über sich selbst gefunden hat

Zehn definierte Fehler; jeder muss GENAU den zuständigen Check rot machen und nach dem
Zurücknehmen wieder grün sein. Angesetzt wird nur an Abschnitt 0 und 1b, weil die nichts
anfassen und nichts reparieren.

Beim ersten echten Lauf: 9 von 9 erkannt — und die Rücknahme von Mutation 9 schlug fehl.
Sabotage und Rücknahme suchten die Schaltfläche beide über „die erste Form MIT Makro";
nach der Sabotage war das eine ANDERE. **Merksatz: Eine Rücknahme darf ihr Ziel nie über
eine Eigenschaft suchen, die die Sabotage selbst verändert hat.**

Daraus wurde zusätzlich die Prüfung „Jede Schaltfläche hat ihr eigenes Makro" — der alte
Check sah nur LEERE Ziele, eine Schaltfläche mit dem Makro einer anderen war unsichtbar.
Mutation 10 weist sie nach.

Weitere Fallen beim Mutationstesten: vorher sicherstellen, dass alles grün ist;
nachweisen, dass die Sabotage überhaupt gegriffen hat; verlangen, dass die übrigen Checks
grün bleiben; benennen, welche Sabotagen der übliche Reparaturweg NICHT heilt.

Statische Gegenprobe vor dem Ausliefern: zeigt jeder `MutErwartet`-Text auf GENAU EINE
`Chk`-Zeile? Dafür Fortsetzungszeilen zusammenfassen und die Literale einzeln durchsuchen.

## Abschnitte des Selbsttests (Ebene 4)

```
0   Erreichbarkeit trotz Option Private Module (Application.Run-Sonde, OnAction-Ziele)
1   Aufbau und Layout
1b  Zustand der Mappe — ALLE Prüfungen ohne Nebenwirkung an einer Stelle
2   Formelneuaufbau darf nichts verändern
3   Wochenplan neu aufbauen        3b  Tabellenende und Zeilen-Schaltflächen
4   Zeilen verschieben             5   Einfügen/Löschen (je Rundlauf)
6   Fixieren und Referenz-Codes    7   Blattschutz
8   Titelblock und PDF             9   Fokus nach den Schaltflächen
10  Übernahme aus einer bisherigen Datei — läuft ALS LETZTER
```

Abschnitt 10 muss dieselben Post-Schritte fahren wie die Schaltfläche
(`Setup_Stoffverteilungsplan` + `UW_Und_Ferien_Generieren`) — sonst prüft er einen Ablauf,
den es nicht gibt.

Der Bericht landet in Mappenordner / `%TEMP%` / Dokumente / `C:\Temp` (erste schreibbare
Stelle); durchgefallene Prüfungen stehen zusätzlich in der Schlussmeldung.

## Ebene 9 — warum sie nicht in die CI kann

Ebene 1–8 prüfen, ob der Code zu seiner eigenen Auslegung passt. Ebene 9 prüft die
Auslegung. Dafür braucht es jemanden, der die Anforderung liest und selbst versteht — ein
Sprachmodell, kein Skript. Ein GitHub-Runner kann das nicht.

Der Ablauf ist deshalb als Arbeitsschritt dokumentiert, mit fertigem Auftragstext:
`Doku/Ebene9_Gegenlese.md`. Ergebnisse werden Issues über
`.github/ISSUE_TEMPLATE/gegenlese-befund.yml`, nicht stille Änderungen.

Beim ersten Einsatz fanden zwei Agenten fünf bestätigte Fehler, die durch alle Ebenen 1–8
gerutscht waren. Die stehen in `Doku/Offene_Punkte.md`.

---

## Ebene W — die Ausgabeschicht (ci_ausgabe.py)

Keine Prüfebene des Excel-Systems, sondern Werkzeug: GitHub wertet nur den Rückgabewert
aus. Das beantwortet „rot oder grün", nicht „wo". Wer das im Log suchen muss, liest bald
nicht mehr nach — und eine Prüfung, die niemand liest, ist keine.

`ci_ausgabe.py` bedient die drei GitHub-Kanäle: **Annotationen** (`::error file=…,line=…`,
erscheinen an der Codezeile), **Job Summary** (Markdown-Tabelle auf der Laufseite) und
**Log-Gruppen**. Außerhalb der CI geben alle Funktionen `None` zurück und drucken nichts —
der lokale Lauf ist unverändert. Die Prüfskripte behalten jeden ihrer `print`-Aufrufe und
bekommen nur eine Zeile dazu.

**Kein `unittest`/`pytest`.** Die Skripte sind keine Unit-Tests, sondern Dateiprüfer, die
auch von Hand als Auslieferungs-Gate laufen (`vbacheck.py mod*.bas`). Ein Test-Runner hätte
diesen Aufruf zerschlagen und einen zweiten Ausführungsweg geschaffen — genau das
Auseinanderlaufen von Arbeitsplatz und Pipeline, gegen das `pruefe_alles.py` gebaut ist.

Ein Fallstrick beim Bau, der hierher gehört: Eine erste Mutation verbog die
Backslash-Ersetzung in `repo_relativ`. Die kann auf Linux gar nicht anschlagen — es gibt
dort keine Backslashes in Pfaden. Der Mutationstest wäre in der CI rot gewesen, ohne dass
an der Prüfung etwas falsch ist. **Merksatz: eine Mutation, die nur auf einem der beiden
Systeme greift, ist keine Mutation.**

---

## Letzter Stand (06.09.2026)

Excel-Selbsttest: **139 bestanden, 0 durchgefallen.** Excel-Mutationstest: **10 von 10.**
`vbacheck --selbsttest`: **9 von 9.** `pruefe_module --selbsttest`: **7 von 7.**
`pruefe_datei --selbsttest`: **10 von 10.** `pruefe_formeln --selbsttest`: **4 von 4.**
Wie viele Prüfungen und Mutationen es jeweils sind, sagt der Lauf selbst — hier steht
bewusst keine Zahl. Jede muss bei jedem Push erneut nachweisen, dass sie noch anschlägt.

`pruefe_alles.py` auf dem Windows-Rechner des Nutzers: **alle 11 Schritte grün.**

### Was die Migration nach Claude Code aufgedeckt hat

Der Stand vom 05.09. behauptete „alle 8 Schritte grün auf dem Rechner des Nutzers". Das
konnte nicht stimmen: `pruefe_module.py --selbsttest` schrieb nach `/tmp`, das es unter
Windows nicht gibt. Der Lauf muss in der Linux-VM stattgefunden haben. **In der CI wäre
das nie aufgefallen** — dort läuft Ubuntu. Eine Prüfung, die nur auf einem der beiden Wege
läuft, unterläuft den erklärten Zweck von `pruefe_alles.py`.

Dahinter steckte ein zweiter, schwererer Fehler: Die Arbeitskopie hieß
`_pruefe_module_test.bas`. Regel 7 vergleicht `Attribute VB_Name` mit dem **Dateinamen** —
unter einem Fantasienamen meldete schon die *unveränderte* Datei `kopf`. Genau diese Regel
erwarten zwei der sieben Mutationen („Kopfzeile fehlt", „Option Explicit fehlt"). Beide
waren seit ihrer Entstehung grün, ohne je etwas nachgewiesen zu haben. Nachgewiesen wurde
das, indem eine saubere Datei unter beiden Namen geprüft wurde: unter dem Fantasienamen
`['kopf']`, unter ihrem echten Namen sauber.

Die Kopie behält jetzt den Namen ihrer Quelldatei, das Verzeichnis kommt von `tempfile`.

**Das ist bereits der dritte Fall in diesem Projekt, in dem eine Prüfung strukturell nicht
anschlagen konnte** — nach der U+FFFD-Suche in Ebene 7 und der Leerzeilen-Prüfung in
Ebene 4. Alle drei fielen erst durch einen Mutationstest auf, keiner durch Nachdenken.
