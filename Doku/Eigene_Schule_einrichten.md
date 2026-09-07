# Eigene Schule einrichten

Das Repository enthält absichtlich keine echten Schuldaten: Die Schulliste ist mit
`Schule 1` und `Schule 2` vorbelegt, die eingebetteten Logos sind neutrale Banner. Diese
Datei beschreibt, wie man seine eigenen einträgt — und was davon ein Update überlebt.

## Was wo liegt

| | wo | überlebt ein Update |
|---|---|---|
| Schulnamen | Blatt „Einstellungen", Spalte der Schulliste | ja, sie stehen in **deiner** Mappe |
| Logos | eingebettet in der Mappe (`wpKopfLogo_<n>`) | nein — eine neue Vorlage bringt die Platzhalter mit |
| Blattschutz-Kennwort | `Makros/modKonfig.bas` | ja, `.gitignore` schließt die Datei aus |
| eigene Suchmuster für Ebene A | `Makros/anonym_muster.local.json` | ja, ebenfalls ausgeschlossen |

Die **Reihenfolge der Schulzeilen bestimmt die Logo-Zuordnung**: die erste Schule bekommt
`wpKopfLogo_1`, die zweite `wpKopfLogo_2`.

## Einmalig einrichten

1. **Schulnamen** im Blatt „Einstellungen" über die vorbelegten `Schule 1` / `Schule 2`
   schreiben. Mehr Schulen: weitere Zeilen darunter — dann braucht es auch entsprechend
   viele eingebettete Logos.

2. **Eigene Logos einsetzen**, in derselben Reihenfolge:

   ```
   cd Makros
   python3 logos_einsetzen.py ../Vorlage/Stoffverteilungsplan_Template.xlsm logo1.png logo2.png
   ```

   Nur PNG, und **breite Banner**: die Platzhalter sind 714×80 und 775×80. Das Skript
   warnt bei einem Verhältnis unter 3 — nachträglich prüft das **nichts**. Der Selbsttest
   misst die Maße der *Form*, und die bleibt beim Tausch unverändert; ein quadratisches
   Logo wird also stillschweigend in den flachen Rahmen gezerrt.

3. **`Makros/modKonfig.bas.vorlage` nach `modKonfig.bas` kopieren** und ein eigenes
   Blattschutz-Kennwort eintragen.

4. **`Makros/anonym_muster.local.json.vorlage` nach `anonym_muster.local.json` kopieren**
   und den eigenen Namen, Schulnamen und das Kennwort als Suchmuster eintragen. Damit
   meldet Prüfebene A, falls je wieder etwas davon in `Vorlage/` landet.

   Ohne diese Datei **bricht `pruefe_anonym.py` lokal mit einem Fehler ab** — genau diese
   Lücke hat schon einmal das Kennwort ins Repository gebracht.

## Nach einem Update

Kennwort, Suchmuster und die Schulnamen im Blatt bleiben unberührt. Neu einzusetzen sind
nur die **Logos**, weil sie in der Mappe stecken:

```
cd Makros
python3 logos_einsetzen.py <neue-mappe.xlsm> logo1.png logo2.png
python3 pruefe_alles.py
```

## Das Kennwort ändern

Nur solange die Mappe **ungeschützt** ist, sonst sperrt man sich aus:

1. „Blattschutz ein / aus" → aus (geht noch mit dem alten Kennwort)
2. erst jetzt `SCHUTZ_PW` in `modKonfig.bas` ändern
3. Module neu importieren, „Einrichtung / Reparatur", Blattschutz wieder ein

## Eine neue Fassung der Vorlage bauen

Wer die Mappen unter `Vorlage/` neu erzeugt, baut sie aus einer **frischen, leeren Mappe**
— nicht aus einer bestehenden. Zwei Gründe, beide belegt:

- Das eingebettete VBA-Projekt enthält sonst weiter die **alte** Fassung geänderter Module.
  Beim Entbranden blieben so die Schulnamen in der Mappe stehen, obwohl die `.bas` sauber
  war.
- Excel schreibt den Speicherort in `xl/workbook.xml`. Liegt die Datei in OneDrive oder
  SharePoint, steht dort eine URL mit Benutzer- und Mandantennamen. **Außerhalb der Cloud
  speichern.**

Und: `modKonfig.bas.vorlage` mit dem Platzhalter-Kennwort importieren, nicht das eigene
Modul. Sonst liegt das Kennwort einkompiliert im Repository.

Der Ablauf im Einzelnen steht unter IMPORTIEREN in `Makros/LIESMICH.txt`.

## Warum die Logos nicht automatisch aus einem Ordner kommen

Das gab es früher und ist bewusst entfallen: `Shape.Copy` mit `Worksheet.Paste` **rastert
das Bild in Anzeigegröße neu** — aus einem 714×80-Banner wurde ein unlesbares 66×55-Bild.
Eine Wiedereinführung über `Shapes.AddPicture` wäre möglich und ist als Issue erfasst.
