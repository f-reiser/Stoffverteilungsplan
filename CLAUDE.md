# Stoffverteilungsplan — Projektanweisungen

Excel-VBA-System für Stoffverteilungspläne an einem bayerischen Gymnasium/FOS.
Der Nutzer ist Mathematiklehrer, kein Vollzeit-Entwickler; er übergibt das Programmieren
vollständig. Er versteht technische Begründungen und will sie hören — aber kurz.

Diese Datei ist der Übergabestand aus einer Cowork-Sitzung vom 31.08. bis 05.09.2026.
Lies bei Projektbeginn zusätzlich `Doku/Offene_Punkte.md` — dort steht, wo es steht.

---

## Die fünf Regeln, deren Bruch schon Schaden angerichtet hat

**1. `.bas`-Dateien sind cp1252 (ANSI) und werden NIE mit Edit/Write bearbeitet.**
Der VBA-Editor liest beim Import ANSI. Das Edit-Werkzeug schreibt UTF-8 zurück und
ersetzt dabei jeden Umlaut durch U+FFFD — der Inhalt ist danach weg. Jede Änderung über
ein Python-Skript mit `io.open(p, encoding='cp1252', newline='')`.

**2. Die bedingte Formatierung wird per VBA GAR NICHT angefasst.**
Sieben der acht Regeln sind x14-Erweiterungsregeln. Der Zugriff auf die klassische
`FormatConditions`-Auflistung hat Excel hart abstürzen lassen.

**3. openpyxl darf diese Mappen NIEMALS speichern.**
Es wirft die x14-Erweiterungen weg und beschädigt `xl/vbaProject.bin`. Lesen ist
unproblematisch. Schreiben nur als XML-Chirurgie im entpackten Zip.

**4. `modKonfig.bas` wird nie ausgeliefert und nie überschrieben.**
Dort steht das Blattschutz-Kennwort des Nutzers **und seine Schulliste**
(`SCHULLISTE`, semikolongetrennt; die Reihenfolge bestimmt die Logo-Zuordnung).
Das Repository ist öffentlich — alles Standortabhängige gehört in dieses Modul oder in
`anonym_muster.local.json`, beide in `.gitignore`. Wie ein Nutzer seine Schule einträgt:
`Doku/Eigene_Schule_einrichten.md`.

**5. Testgetrieben, immer.** Siehe unten.

---

## Arbeitsweise: TDD ist verbindlich

Der Nutzer hat das als Voreinstellung für alle seine Projekte festgelegt; es gibt dazu
einen Skill `test-driven-development` in seinem Konto.

**Zwei Runden, nie eine.** Erst der Test und der Nachweis, dass er ROT ist, dann die
Funktionalität. Bei einem gemeldeten Fehler verschärft: zuerst die Prüfung schreiben, die
den Fehler an SEINER echten Datei rot macht — nicht an einem selbstgebauten
Minimalbeispiel, das bestätigt meist nur die eigene Fehlannahme.

**Wenn die Umgebung zum Rot-Nachweis fehlt** (Excel/VBA läuft in keiner Claude-Umgebung):
den Test ausliefern, ausdrücklich schreiben, *welcher Test mit welcher Meldung fehlschlagen
MUSS*, und ihn den Nachweis führen lassen — erst danach bauen.

**Fremde Gegenlese.** Für alles, was länger trägt: einen unabhängigen Agenten gegenlesen
lassen. Anforderung im Wortlaut des Nutzers mitgeben (nicht als eigene Zusammenfassung —
die trägt die eigene Auslegung schon in sich), eigene Begründungen weglassen, ihn selbst
auslegen lassen, ihn nur MELDEN lassen, Sicherheitsgrad verlangen (sicher / vermutlich /
unklar). Befunde selbst nachprüfen, bevor sie weitergegeben werden. Das hat beim ersten
Einsatz fünf bestätigte Fehler gefunden, die durch alle bestehenden Tests gerutscht waren.

**Warum das hier nicht optional ist:** Test und Code entstehen aus derselben Auslegung.
Ein Check dieses Projekts meldete „keine leeren Zeilen", während 35 dastanden — er hatte
denselben blinden Fleck wie der Fehler, den er finden sollte. Regel daraus: *Wenn der Code
eine Größe mit einer Funktion bestimmt, darf der Test sie nicht mit derselben Funktion
nachprüfen.*

---

## Was wo liegt

```
Stoffverteilungsplan/
├─ Makros/                    ← Quelle. Nur hier wird geschrieben.
│  ├─ mod*.bas                10 Module, cp1252
│  ├─ modKonfig.bas           KENNWORT — in .gitignore, nie ausliefern
│  ├─ modKonfig.bas.vorlage   Platzhalter, damit die Prüfungen ohne das echte Modul laufen
│  ├─ *_Modul.txt             Blattmodul + DieseArbeitsmappe, zum Hineinkopieren
│  ├─ LIESMICH.txt            technische Referenz, 10. Fassung — WICHTIG, ausführlich
│  ├─ pruefe_alles.py         ← EIN Aufruf für alle automatisierbaren Ebenen
│  ├─ vbacheck.py             Ebene 1: statische Prüfung der Module, mit --selbsttest
│  ├─ pruefe_module.py        Ebene 7: Kodierung/Auslieferbarkeit, mit --selbsttest
│  ├─ pruefe_datei.py         Ebenen 2+8: Abnahme fertiger .xlsm, mit --selbsttest
│  ├─ pruefe_formeln.py       Ebene 6: Formelkonsistenz, mit --selbsttest
│  ├─ pruefe_anonym.py        Ebene A: keine Produktivdaten in Vorlage/, mit --selbsttest
│  ├─ anonym_muster.local.json  eigene Namen als Suchmuster — in .gitignore
│  ├─ anonymisiere.py         erzeugt Vorlage/ aus einer produktiven Mappe
│  ├─ logos_einsetzen.py      tauscht die eingebetteten Logos einer Mappe
│  └─ ci_ausgabe.py           Ausgabeschicht für GitHub Actions, mit --selbsttest
├─ Vorlage/                   die einzigen .xlsm im Repo — ERZEUGT, nicht gepflegt
├─ Bilder/                    Schullogos
├─ Diagnose/                  hier legt der Nutzer Testberichte ab (nicht im Repo)
├─ *.xlsm                     Ergebnis, nicht Quelle (nicht im Repo)
├─ .github/workflows/         CI: ruft Makros/pruefe_alles.py bei jedem Push
└─ Doku/                      diese Übergabe
```

**Vor jeder Auslieferung und nach jeder Änderung an den Modulen:**

```
cd Makros && python3 pruefe_alles.py
```

Das ist derselbe Aufruf, den die CI macht. Rückgabewert 0 = grün. Alles andere ist ein
Grund, nicht auszuliefern.

**`Makros/LIESMICH.txt` ist die eigentliche technische Referenz** — Modulaufteilung,
Layout-Erkennung, Update-Strategie, Prüfebenen, Fallstricke, alles mit Begründung.
Vor jeder größeren Änderung lesen.

Die `.bas`-Module sind die **Quelle**, die `.xlsm` das **Ergebnis**. Git versioniert die
Module; fertige Mappen gehören nicht in die Historie (Binärdateien, nicht vergleichbar).

## Was ins Repository darf — und was nicht

**`.gitignore` ist eine Whitelist:** erst wird alles ignoriert, dann gezielt zugelassen.
Grund ist die Arbeitsteilung — erzeugter Quelltext gehört hinein, die Dateien des Nutzers
(fertige Mappen mit echten Unterrichtsdaten, Diagnoseberichte, Sicherungen) nicht. Eine
Blacklist müsste jedes neue Verzeichnis daran denken, sich auszuschließen.

**Daraus folgt eine Pflicht:** Legt der Nutzer neue Verzeichnisse oder Dateien an, muss
vor dem Commit **nachgefragt** werden, ob sie ins Repository sollen. Sonst fehlt still
etwas Wichtiges.

**`Vorlage/` wird erzeugt, nicht gepflegt.** Beide Mappen dort entstehen aus einer
produktiven Datei über `Makros/anonymisiere.py`. Bis 06.09.2026 war `Referenzmappe.xlsm`
eine byteweise Kopie des echten Mathe-Gym-10-Plans — samt Klarname, Schulname und dem
vollständigen OneDrive-Pfad im versteckten Namen `wpPdfOrdner`. `pruefe_anonym.py`
(Ebene A) prüft das bei jedem Push. Von Hand dort hineinschreiben ist ein Fehler.

## Die Module

`modKonfig` (nur `SCHUTZ_PW`, nie anfassen) · `modWochenplan` (Kern: Layout-Erkennung,
Zeilenoperationen, Formeln, Meldungsschleuse) · `modKalender` · `modKopf` (Titelblock,
PDF) · `modSteuerung` (Blatt „Steuerung") · `modSchutz` · `modAnleitung` ·
`modUebernahme` (Datenimport aus einer älteren Mappe) · `modSelbsttest` ·
`modStart` (die vier von Hand startbaren Makros).

Alle außer `modStart` und `modKonfig` tragen `Option Private Module` — das kürzt die
Alt+F8-Liste von 30 auf 4 Einträge und wirkt sich nicht auf `Shape.OnAction` aus
(Abschnitt 0 des Selbsttests weist das bei jedem Lauf nach).

**Es ist nirgends eine Zeilennummer fest verdrahtet.** Alles wird über Beschriftungen
gefunden (`LBL_*`-Konstanten in `modWochenplan`). Der Nutzer verschiebt Dinge im Blatt
„Einstellungen", und das muss weiter funktionieren.

## Die Update-Strategie (wichtig zu verstehen)

Makros lassen sich nicht in fremde Mappen bringen — jeder Weg hängt an einer
Sicherheitseinstellung, die man Kollegen nicht zumuten kann. Deshalb läuft es andersherum:
**nicht neuer Code in die alte Mappe, sondern alte DATEN in die neue Mappe.**
Der Kollege öffnet die neue, leere Datei und drückt „Daten importieren".
`modUebernahme` liest die alte Datei schreibgeschützt aus. Braucht keine
Trust-Center-Einstellung, keine Signatur, kein externes Werkzeug.

---

## Weiterführend

- `Doku/Offene_Punkte.md` — **zuerst lesen.** Aktueller Stand, fünf bestätigte, noch nicht
  reparierte Befunde aus der Gegenlese, Auslegungsfragen für den Nutzer.
- `Doku/Arbeitsweise_und_Regeln.md` — Absprachen mit dem Nutzer, Kodierung, Ausliefern.
- `Doku/Fallstricke.md` — bestätigte Fallen im Excel-Objektmodell und in VBA.
- `Doku/Testebenen.md` — die neun Prüfebenen, der Mutationstest, was er über sich selbst
  gefunden hat. **Regel: jede Prüfung hat eine Mutation, oder sie zählt nicht.**
- `Doku/Ebene9_Gegenlese.md` — der Auftragstext für die fremde Gegenlese, wörtlich
  verwendbar. Die einzige Ebene, die nicht in die CI kann.
- `Doku/Projektwissen.md` — Aufbau der Mappe, XML-Chirurgie, LehrplanPLUS-Recherche,
  Lambacher-Schweizer-Seitenzuordnung.

## Was NICHT in dieser Datei stehen muss

Das persönliche Gedächtnis des Nutzers (Profil, Voreinstellungen) und der Skill
`test-driven-development` hängen an seinem Konto und stehen in jeder Claude-Oberfläche
zur Verfügung. Diese Datei ergänzt sie um das Projektwissen.
