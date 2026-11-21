# =====================================================================
#  BMP-Scanner.ps1
#  Liest den Honeywell-Scanner im Modus "USB Serial" (COM-Port) aus.
#
#  - Medikationsplan (BMP, Text beginnt mit <MP ...):
#      wird als XML-Datei gespeichert, z. B.
#      P:\BMP\Schade_Gerald_17.12.1962_2026-11-21_15-40-12.xml
#      und zusaetzlich in die Zwischenablage gelegt.
#  - Alle anderen (kurzen) Barcodes, z. B. PZN auf Packungen:
#      werden wie bisher in das aktive Fenster "getippt".
#
#  - XML-Dateien in P:\BMP, die aelter als 15 Minuten sind, werden
#      jede Minute nach P:\BMP\alt verschoben.
#
#  Start ueber "BMP-Scanner starten.bat". Fenster darf minimiert bleiben.
# =====================================================================

param(
    [string]$Port      = "",          # leer = Honeywell-Scanner automatisch suchen
    [string]$Ordner    = "P:\BMP",    # Zielordner fuer die XML-Dateien
    [bool]  $KurzcodesTippen = $true, # andere Barcodes ins aktive Fenster tippen
    [bool]  $EnterNachKurzcode = $true, # danach Enter senden (wie vorher im Tastaturmodus)
    [int]   $AltNachMinuten = 15       # XML-Dateien in $Ordner, die aelter sind, nach $Ordner\alt verschieben
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)   # ISO-8859-1, Byte fuer Byte
$ErsatzOrdner = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "BMP"

function Log($text) { Write-Host ("{0:HH:mm:ss}  {1}" -f (Get-Date), $text) }

function Finde-ScannerPort {
    if ($Port -ne "") { return $Port }
    # Honeywell hat die USB-Hersteller-ID 0C2E
    $geraete = Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '\((COM\d+)\)' -and
                       ($_.PNPDeviceID -match 'VID_0C2E' -or $_.Name -match 'Honeywell') }
    foreach ($g in $geraete) {
        if ($g.Name -match '\((COM\d+)\)') { return $Matches[1] }
    }
    return $null
}

function Bereinige-Dateiname([string]$s) {
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $s = $s.Replace([string]$c, "") }
    return ($s.Trim() -replace '\s+', ' ')
}

function Attribut([string]$tag, [string]$name) {
    if ($tag -match ('\s' + $name + '="([^"]*)"')) { return $Matches[1] }
    return ""
}

function Speichere-BMP([string]$xml) {
    # Patientendaten aus <P ... g="Vorname" f="Nachname" b="JJJJMMTT" ...>
    $pTag = ""
    if ($xml -match '<P\s[^>]*>') { $pTag = $Matches[0] }
    $vorname  = Bereinige-Dateiname (Attribut $pTag "g")
    $nachname = Bereinige-Dateiname (Attribut $pTag "f")
    $geb      = Attribut $pTag "b"
    if ($geb -match '^(\d{4})(\d{2})(\d{2})$') { $geb = "$($Matches[3]).$($Matches[2]).$($Matches[1])" }
    $geb = Bereinige-Dateiname $geb
    if ($nachname -eq "") { $nachname = "Unbekannt" }

    # mehrseitiger Plan: <MP ... a="Blatt" z="Blattzahl">
    $mpTag = ""
    if ($xml -match '<MP\s[^>]*>') { $mpTag = $Matches[0] }
    $blatt = Attribut $mpTag "a"; $von = Attribut $mpTag "z"
    $blattTeil = ""
    if ($blatt -ne "" -and $von -ne "" -and $von -ne "1") { $blattTeil = "_Blatt$($blatt)von$($von)" }

    $teile = @($nachname)
    if ($vorname -ne "") { $teile += $vorname }
    if ($geb -ne "")     { $teile += $geb }
    $name = ($teile -join "_") + $blattTeil + "_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xml"

    $ziel = $Ordner
    try {
        if (-not (Test-Path $ziel)) { New-Item -ItemType Directory -Path $ziel -Force | Out-Null }
        $datei = Join-Path $ziel $name
        [System.IO.File]::WriteAllBytes($datei, $Latin1.GetBytes($xml))
    } catch {
        # z. B. Laufwerk P: gerade nicht erreichbar -> lokal speichern
        if (-not (Test-Path $ErsatzOrdner)) { New-Item -ItemType Directory -Path $ErsatzOrdner -Force | Out-Null }
        $datei = Join-Path $ErsatzOrdner $name
        [System.IO.File]::WriteAllBytes($datei, $Latin1.GetBytes($xml))
        Log "ACHTUNG: $Ordner nicht erreichbar, lokal gespeichert!"
        [console]::Beep(400, 600)
    }
    try { Set-Clipboard -Value $xml } catch {}
    [console]::Beep(1200, 120); [console]::Beep(1600, 120)
    Log "Medikationsplan gespeichert: $datei"
}

function Aufraeumen {
    # hoechstens einmal pro Minute
    if (((Get-Date) - $script:letztesAufraeumen).TotalSeconds -lt 60) { return }
    $script:letztesAufraeumen = Get-Date
    try {
        if (-not (Test-Path $Ordner)) { return }
        $grenze = (Get-Date).AddMinutes(-$AltNachMinuten)
        $alte = Get-ChildItem -Path $Ordner -Filter *.xml -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $grenze }
        if (-not $alte) { return }
        $altOrdner = Join-Path $Ordner "alt"
        if (-not (Test-Path $altOrdner)) { New-Item -ItemType Directory -Path $altOrdner -Force | Out-Null }
        foreach ($f in $alte) {
            # laeuft auf mehreren PCs gleichzeitig: Fehler (Datei schon verschoben) einfach ignorieren
            try { Move-Item -LiteralPath $f.FullName -Destination $altOrdner -Force -ErrorAction Stop
                  Log "Nach alt verschoben: $($f.Name)" } catch {}
        }
    } catch {}
}

function Tippe-Text([string]$text) {
    # Sonderzeichen fuer SendKeys maskieren
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $text.ToCharArray()) {
        if ('+^%~(){}[]'.Contains([string]$ch)) { [void]$sb.Append('{' + $ch + '}') }
        else { [void]$sb.Append($ch) }
    }
    if ($EnterNachKurzcode) { [void]$sb.Append('{ENTER}') }
    [System.Windows.Forms.SendKeys]::SendWait($sb.ToString())
}

function Verarbeite([string]$daten) {
    $daten = $daten.Trim([char[]]@(13, 10, 0, 32))
    if ($daten -eq "") { return }
    $start = $daten.IndexOf("<MP")
    if ($start -ge 0) {
        $ende = $daten.LastIndexOf("</MP>")
        if ($ende -gt $start) { $xml = $daten.Substring($start, $ende + 5 - $start) }
        else { $xml = $daten.Substring($start); Log "Warnung: Ende </MP> fehlt, trotzdem gespeichert." }
        # denselben Plan nicht doppelt speichern (z. B. Scanner im Staender)
        if ($xml -eq $script:letzterPlan -and ((Get-Date) - $script:letzterPlanZeit).TotalSeconds -lt 10) {
            Log "Gleicher Plan erneut gelesen - ignoriert."; return
        }
        $script:letzterPlan = $xml; $script:letzterPlanZeit = Get-Date
        Speichere-BMP $xml
    } elseif ($KurzcodesTippen) {
        Tippe-Text $daten
        Log "Barcode getippt: $daten"
    }
}

# ---------------------------- Hauptschleife ----------------------------
$Host.UI.RawUI.WindowTitle = "BMP-Scanner (nicht schliessen)"
$script:letzterPlan = ""; $script:letzterPlanZeit = Get-Date "2000-01-01"
$script:letztesAufraeumen = Get-Date "2000-01-01"
Log "BMP-Scanner gestartet. Zielordner: $Ordner"

while ($true) {
    $com = Finde-ScannerPort
    Aufraeumen
    if (-not $com) { Log "Kein Honeywell-Scanner im COM-Modus gefunden - neuer Versuch in 5 s ..."; Start-Sleep 5; continue }

    $sp = New-Object System.IO.Ports.SerialPort $com, 115200, ([System.IO.Ports.Parity]::None), 8, ([System.IO.Ports.StopBits]::One)
    $sp.Encoding = $Latin1
    $sp.DtrEnable = $true; $sp.RtsEnable = $true
    try {
        $sp.Open()
        Log "Bereit an $com - Medikationsplan einfach scannen."
        $puffer = ""; $letzteDaten = Get-Date
        while ($true) {
            $neu = $sp.ReadExisting()
            if ($neu.Length -gt 0) { $puffer += $neu; $letzteDaten = Get-Date }
            # vollstaendiger Plan da, oder 400 ms Ruhe nach einem Kurzcode
            $fertig = ($puffer.Contains("</MP>")) -or
                      ($puffer.Length -gt 0 -and ((Get-Date) - $letzteDaten).TotalMilliseconds -gt 400)
            if ($fertig) {
                if ($puffer.Contains("</MP>")) {
                    $i = $puffer.IndexOf("</MP>") + 5
                    $teil = $puffer.Substring(0, $i); $puffer = $puffer.Substring($i)
                } else { $teil = $puffer; $puffer = "" }
                try { Verarbeite $teil } catch { Log "Fehler: $($_.Exception.Message)" }
            }
            Aufraeumen
            Start-Sleep -Milliseconds 50
        }
    } catch {
        Log "Verbindung zu $com unterbrochen ($($_.Exception.Message)) - neuer Versuch in 5 s ..."
        Start-Sleep 5
    } finally {
        try { if ($sp.IsOpen) { $sp.Close() } } catch {}
        $sp.Dispose()
    }
}
