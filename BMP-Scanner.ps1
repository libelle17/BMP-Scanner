# =====================================================================
#  BMP-Scanner.ps1
#  Liest den Honeywell-Scanner im Modus "USB Serial" (COM-Port) aus.
#
#  - Medikationsplan (BMP, Text beginnt mit <MP ...):
#      wird als XML-Datei gespeichert, z. B.
#      P:\BMP\Schade_Gerald_17.12.1962_2026-11-21_15-40-12.xml
#      und zusaetzlich in die Zwischenablage gelegt. Ist P: nicht verbunden,
#      ueber \\linux1\daten\Patientendokumente\BMP (derselbe Ordner),
#      erst wenn beides scheitert in Dokumente\BMP.
#  - Alle anderen (kurzen) Barcodes, z. B. PZN auf Packungen:
#      werden wie bisher in das aktive Fenster "getippt".
#
#  - XML-Dateien in P:\BMP, die aelter als 15 Minuten sind, werden
#      jede Minute nach P:\BMP\alt verschoben.
#
#  - In der Taskleiste eigenes Symbol (Data-Matrix-Code) mit Statuspunkt:
#      gruen = bereit, orange = wartet auf den Scanner; Status auch im Titel.
#
#  Start ueber NVerb (Autostart "NVerb.exe bmpscanner"). Fenster darf minimiert bleiben.
# =====================================================================

param(
    [string]$Port      = "",          # leer = Honeywell-Scanner automatisch suchen
    [string]$Ordner    = "P:\BMP",    # Zielordner fuer die XML-Dateien
    [string]$Netzordner = "\\linux1\daten\Patientendokumente\BMP", # derselbe Ordner ohne Laufwerksbuchstaben, falls P: (noch) nicht verbunden ist
    [bool]  $KurzcodesTippen = $true, # andere Barcodes ins aktive Fenster tippen
    [bool]  $EnterNachKurzcode = $true, # danach Enter senden (wie vorher im Tastaturmodus)
    [int]   $AltNachMinuten = 15       # XML-Dateien in $Ordner, die aelter sind, nach $Ordner\alt verschieben
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)   # ISO-8859-1, Byte fuer Byte
$ErsatzOrdner = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "BMP"

function Log($text) { Write-Host ("{0:HH:mm:ss}  {1}" -f (Get-Date), $text) }

# ------------------- Icon und Titel fuer die Taskleiste -------------------
# Eigenes Symbol (Data-Matrix-Code wie auf dem Medikationsplan) mit Statuspunkt:
# gruen = bereit, orange = wartet auf den Scanner. Eigene AppUserModelID, damit
# Windows das Fenster nicht unter dem PowerShell-Symbol zusammenfasst.
# Alles nur Anzeige: geht hier etwas schief, laeuft das Skript normal weiter.
$script:IconOk = $false
$script:Status = ""
$script:Icons = @{}
try {
    Add-Type -AssemblyName System.Drawing
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class BmpFenster {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY { public Guid fmtid; public uint pid; }
    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT { [FieldOffset(0)] public ushort vt; [FieldOffset(8)] public IntPtr p; [FieldOffset(12)] public int p2; }
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore {
        int GetCount(out uint c);
        int GetAt(uint i, out PROPERTYKEY k);
        int GetValue(ref PROPERTYKEY k, out PROPVARIANT v);
        int SetValue(ref PROPERTYKEY k, ref PROPVARIANT v);
        int Commit();
    }
    [DllImport("shell32.dll")] static extern int SHGetPropertyStoreForWindow(IntPtr h, ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyStore ps);
    public static void SetzeAppId(IntPtr h, string id) {
        Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
        IPropertyStore ps;
        if (SHGetPropertyStoreForWindow(h, ref iid, out ps) != 0 || ps == null) return;
        PROPERTYKEY k = new PROPERTYKEY();
        k.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"); k.pid = 5; // PKEY_AppUserModel_ID
        PROPVARIANT v = new PROPVARIANT();
        v.vt = 31; // VT_LPWSTR
        v.p = Marshal.StringToCoTaskMemUni(id);
        ps.SetValue(ref k, ref v);
        ps.Commit();
        Marshal.FreeCoTaskMem(v.p);
        Marshal.ReleaseComObject(ps);
    }
    public static void SetzeIcon(IntPtr h, IntPtr klein, IntPtr gross) {
        SendMessage(h, 0x80, (IntPtr)0, klein);  // WM_SETICON, ICON_SMALL
        SendMessage(h, 0x80, (IntPtr)1, gross);  // WM_SETICON, ICON_BIG
    }
}
"@
    $script:Konsole = [BmpFenster]::GetConsoleWindow()
    if ($script:Konsole -ne [IntPtr]::Zero) {
        [BmpFenster]::SetzeAppId($script:Konsole, "Praxis.BMP-Scanner")
        $script:IconOk = $true
    }
} catch { Log "Hinweis: Symbol fuer die Taskleiste nicht moeglich ($($_.Exception.Message))" }

function Erzeuge-Icon([int]$g, [System.Drawing.Color]$farbe) {
    # Data-Matrix-Muster: durchgehende Kante links und unten, Taktmuster oben und rechts
    $muster = @("1010101010",
                "1001101101",
                "1110010100",
                "1011011011",
                "1100110100",
                "1011100111",
                "1101011000",
                "1001110111",
                "1110100100",
                "1111111111")
    $n = if ($g -le 16) { 7 } else { 10 }
    $m = [math]::Floor(($g - 2) / $n)
    $o = [math]::Floor(($g - $n * $m) / 2)
    $bmp = New-Object System.Drawing.Bitmap $g, $g
    $gr = [System.Drawing.Graphics]::FromImage($bmp)
    $gr.Clear([System.Drawing.Color]::White)
    $dunkel = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(20, 40, 90))
    for ($z = 0; $z -lt $n; $z++) {
        # beim kleinen Muster die letzte Zeile (durchgehende Kante) behalten
        $zeile = if ($z -eq $n - 1) { $muster[9] } else { $muster[$z] }
        for ($s = 0; $s -lt $n; $s++) {
            $an = ($s -eq 0) -or ($zeile[$s] -eq [char]'1')
            if ($s -eq $n - 1 -and $z -lt $n - 1) { $an = ($z % 2 -eq 1) }   # Taktmuster rechts
            if ($z -eq 0) { $an = ($s % 2 -eq 0) }                           # Taktmuster oben
            if ($an) { $gr.FillRectangle($dunkel, $o + $s * $m, $o + $z * $m, $m, $m) }
        }
    }
    # Statuspunkt unten rechts mit weissem Rand
    $d = [math]::Round($g * 0.5)
    $gr.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $gr.FillEllipse([System.Drawing.Brushes]::White, $g - $d - 1, $g - $d - 1, $d + 1, $d + 1)
    $punkt = New-Object System.Drawing.SolidBrush $farbe
    $rand = [math]::Max(1, [math]::Round($g / 16))
    $gr.FillEllipse($punkt, $g - $d - 1 + $rand, $g - $d - 1 + $rand, $d + 1 - 2 * $rand, $d + 1 - 2 * $rand)
    $gr.Dispose(); $dunkel.Dispose(); $punkt.Dispose()
    return $bmp.GetHicon()
}

function Setze-Status([string]$status, [string]$text) {
    # nur bei Aenderung, die Hauptschleife ruft das oft auf
    if ($status -eq $script:Status) { return }
    $script:Status = $status
    try { $Host.UI.RawUI.WindowTitle = "BMP-Scanner (nicht schliessen) - $text" } catch {}
    if (-not $script:IconOk) { return }
    try {
        if (-not $script:Icons.ContainsKey($status)) {
            $farbe = if ($status -eq "bereit") { [System.Drawing.Color]::FromArgb(30, 160, 60) }
                     else { [System.Drawing.Color]::FromArgb(235, 140, 0) }
            $script:Icons[$status] = @((Erzeuge-Icon 16 $farbe), (Erzeuge-Icon 32 $farbe))
        }
        [BmpFenster]::SetzeIcon($script:Konsole, $script:Icons[$status][0], $script:Icons[$status][1])
    } catch {}
}

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

    # erst P:\BMP, dann derselbe Ordner ueber \\linux1 (P: evtl. nicht verbunden), zuletzt lokal
    $datei = $null
    foreach ($ziel in @($Ordner, $Netzordner)) {
        if ($datei -or [string]::IsNullOrEmpty($ziel)) { continue }
        try {
            if (-not (Test-Path $ziel)) { New-Item -ItemType Directory -Path $ziel -Force | Out-Null }
            $d = Join-Path $ziel $name
            [System.IO.File]::WriteAllBytes($d, $Latin1.GetBytes($xml))
            $datei = $d
            if ($ziel -ne $Ordner) { Log "Hinweis: $Ordner nicht erreichbar, ueber $ziel gespeichert." }
        } catch {}
    }
    if (-not $datei) {
        if (-not (Test-Path $ErsatzOrdner)) { New-Item -ItemType Directory -Path $ErsatzOrdner -Force | Out-Null }
        $datei = Join-Path $ErsatzOrdner $name
        [System.IO.File]::WriteAllBytes($datei, $Latin1.GetBytes($xml))
        Log "ACHTUNG: $Ordner und $Netzordner nicht erreichbar, lokal gespeichert!"
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
        # P:\BMP, sonst derselbe Ordner ueber \\linux1
        $basis = @($Ordner, $Netzordner) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
        if (-not $basis) { return }
        $grenze = (Get-Date).AddMinutes(-$AltNachMinuten)
        $alte = Get-ChildItem -Path $basis -Filter *.xml -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $grenze }
        if (-not $alte) { return }
        $altOrdner = Join-Path $basis "alt"
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
Log "BMP-Scanner gestartet. Zielordner: $Ordner (sonst $Netzordner)"

while ($true) {
    $com = Finde-ScannerPort
    Aufraeumen
    if (-not $com) { Setze-Status "warte" "wartet auf Scanner"; Log "Kein Honeywell-Scanner im COM-Modus gefunden - neuer Versuch in 5 s ..."; Start-Sleep 5; continue }

    $sp = New-Object System.IO.Ports.SerialPort $com, 115200, ([System.IO.Ports.Parity]::None), 8, ([System.IO.Ports.StopBits]::One)
    $sp.Encoding = $Latin1
    $sp.DtrEnable = $true; $sp.RtsEnable = $true
    try {
        $sp.Open()
        Setze-Status "bereit" "bereit an $com"
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
        Setze-Status "warte" "Verbindung zu $com unterbrochen"
        Log "Verbindung zu $com unterbrochen ($($_.Exception.Message)) - neuer Versuch in 5 s ..."
        Start-Sleep 5
    } finally {
        try { if ($sp.IsOpen) { $sp.Close() } } catch {}
        $sp.Dispose()
    }
}
