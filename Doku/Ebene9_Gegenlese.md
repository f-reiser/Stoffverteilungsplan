# Ebene 9 — Fremde Gegenlese

**Diese Ebene kann nicht in die CI.** Sie braucht ein Sprachmodell, das die Anforderung
liest und selbst auslegt. Das ist keine Bequemlichkeit, sondern der Zweck: Ebene 1–8
prüfen, ob der Code zu seiner eigenen Auslegung passt. Ebene 9 prüft die Auslegung.

Damit sie trotzdem nicht wieder zur Gewohnheit verkommt, steht sie hier als Ablauf mit
festem Wortlaut. Sie ist kein Test, sondern ein **Arbeitsschritt vor jeder größeren
Auslieferung** — und ihr Ergebnis sind Issues, keine stillen Änderungen.

---

## Warum überhaupt

Test und Code entstehen aus derselben Auslegung. Wenn eine Instanz die Anforderung liest,
den Code schreibt und den Test schreibt, trägt der Test denselben blinden Fleck wie der
Code. Belegt in diesem Projekt:

- Der Check „Keine leeren Zeilen am Ende des Plans" meldete **ok**, während 35 leere Zeilen
  unter der Tabelle standen — er begann bei `PlanLastRow` und suchte nach oben, die Zeilen
  lagen darunter.
- Der Check „Summenformel zeigt auf die Datenzeilen" war grün, obwohl nichts sie neu gebaut
  hatte. Die alte Formel hatte zufällig überlebt.

Beim ersten Einsatz von Ebene 9 fanden zwei unabhängige Agenten **fünf bestätigte Fehler**,
die durch sämtliche bestehenden Tests der Ebenen 1–8 gerutscht waren.

---

## Die fünf Regeln des Auftrags

**1. Die Anforderung im Wortlaut des Nutzers mitgeben — nicht als eigene Zusammenfassung.**
Eine Zusammenfassung trägt die eigene Auslegung schon in sich. Genau die soll geprüft
werden. In diesem Projekt heißt das: `Makros/LIESMICH.txt` mitgeben, nicht erklären.

**2. Eigene Begründungen weglassen.** Kein „das ist so, weil …". Der fremde Agent soll das
Verhalten aus dem Code ableiten und selbst beurteilen, ob es zur Anforderung passt.

**3. Nur MELDEN lassen, nichts ändern.** Ein Agent, der repariert, verschiebt den blinden
Fleck nur um eine Instanz. Das Ergebnis ist ein Befund, den ein Mensch entscheidet.

**4. Sicherheitsgrad verlangen: sicher / vermutlich / unklar.** Ohne das kommen zwanzig
gleich laute Befunde zurück, und das Sortieren kostet mehr Zeit als die Prüfung gespart hat.

**5. Jeden Befund selbst im Code nachprüfen, bevor er weitergegeben wird.** Beim ersten
Einsatz waren fünf von neun Befunden echt. Ungeprüft weitergereichte Befunde sind
Rauschen.

---

## Der Auftragstext

Wörtlich verwendbar. `<…>` ersetzen.

```
Du prüfst fremden Code gegen eine fremde Anforderung. Du kennst weder den Autor
noch seine Überlegungen, und das ist Absicht.

ANFORDERUNG
    Im Anhang: <LIESMICH.txt> — der Wortlaut, nach dem gebaut wurde.
    Lege ihn selbst aus. Frage nicht nach, was gemeint ist; wenn eine Stelle
    mehrdeutig ist, ist genau das ein Befund.

CODE
    Im Anhang: <die betroffenen Module und die zugehörigen Tests>.

AUFTRAG
    Prüfe, ob der Code und seine Tests die Anforderung erfüllen. Achte besonders auf:
    - Tests, die eine Größe mit derselben Funktion nachprüfen, mit der der Code sie
      bestimmt — solche Tests bestätigen nur, dass der Code zu sich selbst passt.
    - Prüfungen, die aus struktureller Ursache gar nicht anschlagen können.
    - Zusagen der Anforderung, für die es keine Prüfung gibt.

    ÄNDERE NICHTS. Melde.

FORM JEDES BEFUNDS
    Titel:            eine Zeile
    Stelle:           Datei und Prozedur
    Was ich sehe:     das beobachtete Verhalten, ohne Deutung
    Warum das ein Problem ist: Bezug auf die Anforderung, mit Zitat
    Sicherheitsgrad:  sicher | vermutlich | unklar
                      sicher     = im Code nachweisbar
                      vermutlich = plausibel, aber vom Laufzeitverhalten abhängig
                      unklar     = die Anforderung gibt es nicht eindeutig her
```

---

## Wann

- Vor jeder Auslieferung, die mehr als eine Prozedur berührt.
- Immer, wenn eine neue Prüfebene oder ein neuer Test entsteht — der Test ist dann selbst
  der Prüfgegenstand.
- Nach jedem gemeldeten Fehler, den keine bestehende Prüfung gefunden hat. Dann lautet die
  Zusatzfrage: *Welche Prüfung hätte das finden müssen, und warum hat sie es nicht?*

Zwei Agenten sind besser als einer; sie überschneiden sich weniger, als man erwartet.

## Wohin mit dem Ergebnis

Jeder bestätigte Befund wird ein Issue über die Vorlage
`.github/ISSUE_TEMPLATE/gegenlese-befund.yml`. Nicht bestätigte Befunde werden **nicht**
stillschweigend verworfen, sondern mit einer Zeile Begründung geschlossen — sonst kommt
derselbe Befund bei der nächsten Gegenlese wieder.
