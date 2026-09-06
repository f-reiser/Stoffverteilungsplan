<#
  Stoffverteilungsplan - Makros in vorhandene Mappen übertragen
  =============================================================
  Kopiert das VBA-Projekt aus einer Vorlage-Mappe in alle anderen
  .xlsm-Dateien des Ordners. Die Daten der Zieldateien bleiben dabei
  vollständig unberührt - ausgetauscht wird nur der Programmteil.

  Hintergrund: In einer .xlsm liegt der gesamte VBA-Code als EIN
  Baustein (xl/vbaProject.bin). Wird der ersetzt, sind alle Module und
  auch der Code hinter "Tabelle1" und "DieseArbeitsmappe" auf einen
  Schlag aktuell. Dafür braucht es weder Excel noch die Einstellung
  "Zugriff auf das VBA-Projektobjektmodell vertrauen".

  Bedienung:  Makros_verteilen.cmd doppelklicken.

  Schalter:
     -Probelauf    zeigt nur, was passieren würde
     -Ja           ohne Rückfrage durchlaufen
     -Vorlage      andere Vorlagedatei angeben
#>

[CmdletBinding()]
param(
    [string] $Vorlage,
    [string] $Ordner,
    [switch] $Probelauf,
    [switch] $Ja
)

$ErrorActionPreference = 'Stop'
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

if (-not $Ordner) { $Ordner = $PSScriptRoot }
if (-not $Ordner) { $Ordner = (Get-Location).Path }

$SICHERUNG = 'Sicherung'

function Schreib([string]$t, [string]$farbe = 'Gray') {
    Write-Host $t -ForegroundColor $farbe
}

function Lies-VbaProjekt([string]$pfad) {
    $zip = [IO.Compression.ZipFile]::OpenRead($pfad)
    try {
        $e = $zip.GetEntry('xl/vbaProject.bin')
        if (-not $e) { return $null }
        $ms = New-Object IO.MemoryStream
        $s = $e.Open()
        try { $s.CopyTo($ms) } finally { $s.Dispose() }
        return $ms.ToArray()
    } finally { $zip.Dispose() }
}

function Ist-Stoffverteilungsplan([string]$pfad) {
    # Schutz davor, das VBA-Projekt in eine wildfremde Mappe zu kippen:
    # es muss wenigstens die drei tragenden Blätter geben.
    $zip = [IO.Compression.ZipFile]::OpenRead($pfad)
    try {
        $e = $zip.GetEntry('xl/workbook.xml')
        if (-not $e) { return $false }
        $s = $e.Open()
        try {
            $r = New-Object IO.StreamReader($s)
            $xml = $r.ReadToEnd()
        } finally { $s.Dispose() }
        foreach ($blatt in @('Wochenplan', 'Lernbereiche', 'Einstellungen')) {
            if ($xml -notmatch ('name="' + $blatt + '"')) { return $false }
        }
        return $true
    } finally { $zip.Dispose() }
}

function Ist-Frei([string]$pfad) {
    # Excel legt neben einer geöffneten Mappe eine Sperrdatei ~$... an.
    $sperre = Join-Path ([IO.Path]::GetDirectoryName($pfad)) ('~$' + [IO.Path]::GetFileName($pfad))
    if (Test-Path -LiteralPath $sperre) { return $false }
    try {
        $h = [IO.File]::Open($pfad, 'Open', 'ReadWrite', 'None')
        $h.Close()
        return $true
    } catch { return $false }
}

# ---------------------------------------------------------------- Vorlage
if (-not $Vorlage) {
    $kandidaten = @(Get-ChildItem -LiteralPath $Ordner -Filter '*.xlsm' -File |
                    Where-Object { $_.Name -like '*Template*' -and $_.Name -notlike '~$*' })
    if ($kandidaten.Count -eq 1) {
        $Vorlage = $kandidaten[0].FullName
    } elseif ($kandidaten.Count -eq 0) {
        Schreib "Keine Vorlage gefunden." 'Red'
        Schreib "Es braucht eine .xlsm mit 'Template' im Namen, oder den Aufruf" 'Red'
        Schreib "mit  -Vorlage <Pfad>." 'Red'
        exit 1
    } else {
        Schreib "Mehrere mögliche Vorlagen - bitte mit -Vorlage <Pfad> eine angeben:" 'Red'
        $kandidaten | ForEach-Object { Schreib ("   " + $_.Name) }
        exit 1
    }
}
if (-not (Test-Path -LiteralPath $Vorlage)) {
    Schreib "Vorlage nicht gefunden: $Vorlage" 'Red'; exit 1
}
$VorlagePfad = (Resolve-Path -LiteralPath $Vorlage).Path

$vba = Lies-VbaProjekt $VorlagePfad
if (-not $vba) {
    Schreib "Die Vorlage enthält kein VBA-Projekt." 'Red'
    Schreib "Ist es wirklich eine .xlsm, in der die Module schon importiert sind?" 'Red'
    exit 1
}

Schreib ""
Schreib "  Stoffverteilungsplan - Makros verteilen" 'White'
Schreib "  ---------------------------------------" 'White'
Schreib ""
Schreib ("  Vorlage : {0}" -f [IO.Path]::GetFileName($VorlagePfad))
Schreib ("  Makros  : {0:N0} Bytes" -f $vba.Length)
Schreib ("  Ordner  : {0}" -f $Ordner)
Schreib ""

# ----------------------------------------------------- Zieldateien holen
# -File und kein -Recurse: der Unterordner mit den Sicherungskopien
# bleibt dadurch von selbst aussen vor.
$ziele = @(Get-ChildItem -LiteralPath $Ordner -Filter '*.xlsm' -File |
           Where-Object { $_.FullName -ne $VorlagePfad -and $_.Name -notlike '~$*' })

if ($ziele.Count -eq 0) {
    Schreib "  Keine weiteren .xlsm-Dateien in diesem Ordner." 'Yellow'
    exit 0
}

Schreib "  Diese Mappen bekommen die Makros der Vorlage:" 'White'
foreach ($f in $ziele) {
    $frei = Ist-Frei $f.FullName
    $hatVba = $null -ne (Lies-VbaProjekt $f.FullName)
    if (-not $frei)        { Schreib ("     {0}   (in Excel geöffnet - wird übersprungen)" -f $f.Name) 'Yellow' }
    elseif (-not $hatVba)  { Schreib ("     {0}   (keine Makro-Mappe - wird übersprungen)" -f $f.Name) 'Yellow' }
    else                   { Schreib ("     {0}" -f $f.Name) 'Cyan' }
}
Schreib ""

if ($Probelauf) { Schreib "  Probelauf - es wurde nichts geändert." 'White'; exit 0 }

if (-not $Ja) {
    Schreib "  Von jeder Datei wird vorher eine Sicherungskopie angelegt" 'DarkGray'
    Schreib ("  (Unterordner '{0}')." -f $SICHERUNG) 'DarkGray'
    Schreib ""
    $antwort = Read-Host "  Fortfahren? [j/N]"
    if ($antwort -notmatch '^[jJyY]') { Schreib "  Abgebrochen." 'Yellow'; exit 0 }
    Schreib ""
}

$sicherungsOrdner = Join-Path $Ordner $SICHERUNG
$ok = 0; $weg = 0

foreach ($f in $ziele) {
    if (-not (Ist-Frei $f.FullName)) {
        Schreib ("  übersprungen  {0}" -f $f.Name) 'Yellow'; $weg++; continue
    }
    if ($null -eq (Lies-VbaProjekt $f.FullName)) {
        Schreib ("  übersprungen  {0}" -f $f.Name) 'Yellow'; $weg++; continue
    }

    if (-not (Test-Path -LiteralPath $sicherungsOrdner)) {
        New-Item -ItemType Directory -Path $sicherungsOrdner | Out-Null
    }
    $stempel = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $sicherung = Join-Path $sicherungsOrdner `
                 ([IO.Path]::GetFileNameWithoutExtension($f.Name) + " - $stempel.xlsm")
    Copy-Item -LiteralPath $f.FullName -Destination $sicherung -Force

    # Archiv vollständig neu schreiben und dabei nur den VBA-Baustein
    # austauschen. Bewusst nicht im Update-Modus: so bleibt die
    # Reihenfolge der Einträge erhalten und [Content_Types].xml steht
    # wie vorgeschrieben an erster Stelle.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString() + '.xlsm')
    $quelle = [IO.Compression.ZipFile]::OpenRead($f.FullName)
    try {
        $neu = [IO.Compression.ZipFile]::Open($tmp, 'Create')
        try {
            $reihenfolge = @($quelle.Entries | Where-Object { $_.FullName -eq '[Content_Types].xml' }) +
                           @($quelle.Entries | Where-Object { $_.FullName -ne '[Content_Types].xml' })
            foreach ($e in $reihenfolge) {
                $ziel = $neu.CreateEntry($e.FullName, [IO.Compression.CompressionLevel]::Optimal)
                $aus = $ziel.Open()
                try {
                    if ($e.FullName -eq 'xl/vbaProject.bin') {
                        $aus.Write($vba, 0, $vba.Length)
                    } else {
                        $ein = $e.Open()
                        try { $ein.CopyTo($aus) } finally { $ein.Dispose() }
                    }
                } finally { $aus.Dispose() }
            }
        } finally { $neu.Dispose() }
    } finally { $quelle.Dispose() }

    Move-Item -LiteralPath $tmp -Destination $f.FullName -Force
    Schreib ("  aktualisiert  {0}" -f $f.Name) 'Green'
    $ok++
}

Schreib ""
Schreib ("  Fertig: {0} aktualisiert, {1} übersprungen." -f $ok, $weg) 'White'
if ($ok -gt 0) {
    Schreib ""
    Schreib "  Noch zu tun, in JEDER aktualisierten Mappe einmal:" 'White'
    Schreib "     öffnen  ->  Blatt 'Steuerung'  ->  'Einrichtung / Reparatur'" 'White'
    Schreib ""
    Schreib ("  Die Sicherungskopien liegen im Unterordner '{0}' und können" -f $SICHERUNG) 'DarkGray'
    Schreib "  gelöscht werden, sobald alles läuft." 'DarkGray'
}
