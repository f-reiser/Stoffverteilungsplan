# Eigene Schule einrichten

Das Repository enthält absichtlich keine echten Schuldaten: Schulnamen stehen als
Platzhalter, die eingebetteten Logos sind neutrale Banner. Diese Datei beschreibt, wie man
seine eigenen einträgt — und vor allem, wie man das **nach jedem Update** ohne Handarbeit
wiederherstellt.

## Das Prinzip

Alles Standortabhängige liegt in Dateien, die ein Update nie anfasst:

| Datei | Inhalt | im Repository? |
|---|---|---|
| `Makros/modKonfig.bas` | Blattschutz-Kennwort, Schulliste | nein (`.gitignore`) |
| `Makros/anonym_muster.local.json` | eigene Namen als Suchmuster für Ebene A | nein (`.gitignore`) |
| eigene Logo-Dateien | die Bilder selbst | nein |

Zu jeder gibt es eine `.vorlage` im Repository, an der man sieht, was hineingehört. Ein
`git pull` überschreibt keine davon.

## Einmalig einrichten

1. **`Makros/modKonfig.bas.vorlage` nach `Makros/modKonfig.bas` kopieren.** Dort das
   eigene Blattschutz-Kennwort eintragen und die Schulliste:

   ```vba
   Public Const SCHULLISTE As String = "Meine Schule;Meine zweite Schule"
   ```

   Die **Reihenfolge bestimmt die Logo-Zuordnung**: die erste Schule bekommt
   `wpKopfLogo_1`, die zweite `wpKopfLogo_2`.

2. **Eigene Logos einsetzen** — je Schule eines, in derselben Reihenfolge:

   ```
   cd Makros
   python3 logos_einsetzen.py ../Vorlage/Stoffverteilungsplan_Template.xlsm logo1.png logo2.png
   ```

   Die Bilder sollten **breite Banner** sein. Der Selbsttest verlangt ein Verhältnis
   Breite/Höhe über 3; die Platzhalter sind 714×80 und 775×80. Ein quadratisches Logo
   macht den Test rot und sieht im Titelblock gequetscht aus.

3. **`Makros/anonym_muster.local.json.vorlage` nach `anonym_muster.local.json` kopieren**
   und den eigenen Namen und Schulnamen eintragen. Damit meldet Prüfebene A, falls je
   wieder produktive Daten in `Vorlage/` landen. Ohne diese Datei prüft sie nur die
   allgemeinen Spuren und sagt das beim Lauf dazu.

## Nach einem Update

`modKonfig.bas` und die Musterdatei bleiben unberührt — die Schulliste ist also sofort
wieder da. Neu einzusetzen sind nur die **Logos**, weil sie in der Mappe stecken und die
neue Fassung die Platzhalter mitbringt:

```
cd Makros
python3 logos_einsetzen.py <neue-mappe.xlsm> logo1.png logo2.png
python3 pruefe_alles.py
```

Ein Aufruf statt Handarbeit in Excel. Wer die eigenen Logos griffbereit halten will, legt
sie in einen Ordner, den `.gitignore` ohnehin ausschließt — etwa `Bilder/`.

## Warum die Logos nicht automatisch aus einem Ordner kommen

Das gab es früher und ist bewusst entfallen: Der Weg über `Shape.Copy` und `Worksheet.Paste`
**rastert das Bild in Anzeigegröße neu** — aus einem 714×80-Banner wurde ein unlesbares
66×55-Bild. Seitdem sind die Logos fest eingebettet, und der Austausch passiert außerhalb
von Excel auf der Datei.

Eine Wiedereinführung — Logos beim „Einrichtung / Reparatur" automatisch aus einem
konfigurierten Ordner einsetzen — wäre möglich (`Shapes.AddPicture` rastert nicht), ist
aber eine eigene Änderung mit eigenem Test. Siehe das zugehörige Issue.
