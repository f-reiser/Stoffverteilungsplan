# Stoffverteilungsplan

Excel-VBA-System für Stoffverteilungspläne an bayerischen Gymnasien und Fachoberschulen.
Eine Mappe je Fach und Jahrgangsstufe: Unterrichtswochen und Ferien werden aus dem
Schuljahreskalender erzeugt, Lernbereiche mit Zeitrichtwerten gegengerechnet, der Plan als
PDF ausgegeben.

## Aufbau

Die **`.bas`-Module sind die Quelle**, die `.xlsm` das Ergebnis. Versioniert werden die
Module; fertige Mappen sind Binärdateien und gehören nicht in die Historie. Ausgenommen
ist `Vorlage/` mit genau zwei Referenzmappen, damit die Prüfungen etwas zu prüfen haben.

```
Makros/     VBA-Module und die Prüfskripte
Vorlage/    leere Vorlage und eine gefüllte Referenzmappe (anonymisiert)
Doku/       Arbeitsweise, Fallstricke, Prüfebenen, offene Punkte
```

## Prüfen

```
cd Makros && python3 pruefe_alles.py
```

Elf Schritte, 61 Mutationstests. Derselbe Aufruf läuft in der CI bei jedem Push.
Rückgabewert 0 heißt grün. Was sich nicht automatisieren lässt — der Selbsttest in echtem
Excel und die fremde Gegenlese — steht in `Doku/Testebenen.md`.

Der Grundsatz dahinter: **jede Prüfung hat einen Mutationstest, oder sie zählt nicht.**
Eine Prüfung ohne den Nachweis, dass sie überhaupt anschlagen kann, ist eine Behauptung —
dieses Projekt hatte drei davon, und alle drei fielen erst durch ihren Mutationstest auf.

## Für die eigene Schule

Das Repository enthält keine echten Schuldaten: Schulnamen sind Platzhalter, die
eingebetteten Logos neutrale Banner. Wie man seine eigenen einträgt und nach einem Update
mit einem Aufruf wiederherstellt, steht in
[`Doku/Eigene_Schule_einrichten.md`](Doku/Eigene_Schule_einrichten.md).

## Mitarbeit

Anforderungen, Fehler und Rückfragen laufen über Issues, nicht über einen Chat — sie sind
dort durchsuchbar, verlinkbar und für alle sichtbar. Änderungen kommen über Feature-Branch
und Pull Request nach `main`.

## Lizenz

[MIT](LICENSE)
