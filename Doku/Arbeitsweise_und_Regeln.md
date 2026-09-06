# Arbeitsweise und harte Regeln

Absprachen mit dem Nutzer und Regeln, deren Verletzung schon Schaden angerichtet hat.
Stand 05.09.2026.

## Absprachen mit dem Nutzer (wörtlich, gelten weiter)

> „Verändere Daten in dem Lokalen Ordner nur, wenn ich es explizit freigebe. Die Freigabe
> gilt nur für die von mir genannten Dateien, alle anderen Dateien bleiben von dir
> unberührt. Standardmäßig überschreibst du bitte keine meiner Dateien, sondern erstellst
> eine neue Kopie mit dem Postfix _Claude vor der Dateiendung."

> „Nutze außerdem zum Lesen nur die Dateien in dem lokalen Ordner, die ich dir explizit
> nenne."

Freigegeben zum Lesen **und Schreiben**: der Unterordner `Makros`
(„ich hab bereits ein Backup gezogen, es kann also nichts kaputt gehen").
Freigegeben zum Schreiben: `Stoffverteilungsplan_Template2.xlsm` (historisch).
Die Dateien `... - v1.xlsm` / `... - old.xlsm` sind seine eigenen Archive — nur lesen.
Für die Migration hat er das `_Claude`-Postfix einmalig aufgehoben.

Der Nutzer ist Mathematiklehrer, kein Vollzeit-Entwickler. Er hat Erfahrung mit LaTeX,
Batch und aus früherer Zeit mit richtiger Programmierung — er versteht technische
Begründungen und will sie auch hören, aber er hat wenig Zeit. Ergebnisse liefern, die
Begründung dazu, keine Vorlesung.

## TDD ist Pflicht, nicht Kür

Der Nutzer hat testgetriebenes Vorgehen als Voreinstellung für ALLE seine Projekte
festgelegt. Es gibt dazu einen Skill `test-driven-development` in seinem Konto.

Kern: **Zwei Runden, nie eine.** Erst der Test und der Nachweis, dass er ROT ist, dann
die Funktionalität. Bei einem gemeldeten Fehler: erst die Prüfung schreiben, die den
Fehler an SEINER echten Datei rot macht — nicht an einem selbstgebauten Minimalbeispiel.

Wenn die Umgebung zum Rot-Nachweis fehlt (Excel/VBA läuft hier nicht), ausdrücklich
schreiben, **welcher Test mit welcher Meldung fehlschlagen MUSS**, und ihn den Nachweis
führen lassen, bevor gebaut wird.

Zusatz, den er selbst vorgeschlagen hat und der sich sofort bezahlt gemacht hat: für
alles, was länger trägt, **einen unabhängigen Agenten gegenlesen lassen**. Anforderung im
Wortlaut mitgeben, eigene Begründungen weglassen, den Agenten selbst auslegen lassen, ihn
nur MELDEN lassen (nichts ändern), Sicherheitsgrad verlangen (sicher / vermutlich /
unklar). Befunde selbst nachprüfen, bevor sie weitergegeben werden.

## Kodierung — der teuerste Fehler des Projekts

Die `.bas`-Dateien MÜSSEN als **cp1252 (ANSI)** auf der Platte liegen. Der VBA-Editor
liest beim Import ANSI; eine UTF-8-Datei kommt als Hieroglyphen an.

**Die cp1252-Dateien NIEMALS mit dem Edit- oder Write-Werkzeug bearbeiten.** Das Edit-Tool
schreibt UTF-8 zurück und ersetzt dabei jeden Umlaut durch U+FFFD — der Inhalt ist
danach unwiederbringlich weg. `modAnleitung.bas` wurde nur gerettet, weil die
Vorgängerfassung noch beim Nutzer lag.

Jede Änderung an einer `.bas`/`.txt` über ein kleines Python-Skript:

```python
t = io.open(p, encoding='cp1252', newline='').read()
assert t.count(alt) == 1
t = t.replace(alt, neu, 1)
io.open(p, 'w', encoding='cp1252', newline='').write(t)
```

Nach dem Ausliefern IMMER prüfen: kein U+FFFD, keine UTF-8-Reste (`b'\xc3\xa4'` &co.),
keine doppelten Backslashes in Zeichenketten, keine CRLF-Mischung. `grep` in der Linux-VM
findet in cp1252-Dateien nichts (ungültiges UTF-8) — Python nehmen.

## Ausliefern

`.bas` ins Verzeichnis `Makros`, der Nutzer importiert sie im VBA-Editor
(Alt+F11 → Datei → Datei importieren). Blattmodul- und Arbeitsmappencode als `.txt`
zum Hineinkopieren.

**`modKonfig.bas` wird NIE ausgeliefert.** Dort steht sein Blattschutz-Kennwort; jedes
Update würde es überschreiben. Genau dafür wurde es aus `modSchutz` herausgelöst.

Vor jeder Auslieferung: `python3 vbacheck.py mod*.bas` muss sauber durchlaufen.
Nach jeder Auslieferung: `python3 pruefe_datei.py ../*.xlsm` über die fertigen Mappen.

## Der Nutzer legt Berichte im Unterordner `Diagnose` ab

Selbsttest- und Mutationsberichte findet man dort, wenn er sie schickt.
