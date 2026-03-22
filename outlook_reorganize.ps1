param([switch]$Live)
$DryRun = -not $Live

try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch { Write-Host "FEHLER: Outlook ist nicht geoeffnet." -ForegroundColor Red; exit 1 }

$namespace = $outlook.GetNamespace("MAPI")
$inbox     = $namespace.GetDefaultFolder(6)
$st        = @{ created=0; moved=0; merged=0; classified=0; deleted=0; skipped=0 }

$ae=[char]0x00E4; $oe=[char]0x00F6; $ue=[char]0x00FC
$Ae=[char]0x00C4; $Oe=[char]0x00D6; $Ue=[char]0x00DC; $ss=[char]0x00DF

function Log($msg, $color="White") { Write-Host $msg -ForegroundColor $color }

function Get-Folder($path) {
    $cur = $inbox
    foreach ($p in ($path -split "/")) {
        $f = $null; try { $f = $cur.Folders.Item($p) } catch {}
        if ($null -eq $f) { return $null }
        $cur = $f
    }
    return $cur
}

function Ensure-Folder($path) {
    $parts = $path -split "/"
    $cur = $inbox
    foreach ($p in $parts) {
        $f = $null; try { $f = $cur.Folders.Item($p) } catch {}
        if ($null -eq $f) {
            Log ("    [NEU]     $path") "Green"
            $st.created++
            if (-not $DryRun) { $f = $cur.Folders.Add($p) } else { return $null }
        }
        $cur = $f
    }
    return $cur
}

function Move-Folder($src, $destPath, $newName=$null) {
    $f = $null; try { $f = $inbox.Folders.Item($src) } catch {}
    if ($null -eq $f) { Log ("    [SKIP]    $src nicht gefunden") "DarkYellow"; $st.skipped++; return }
    $dest = Get-Folder $destPath
    if ($null -eq $dest) { Log ("    [FEHLER]  Ziel nicht gefunden: $destPath") "Red"; return }
    $n = if ($newName) { $newName } else { $src }
    Log ("    [{0,5} Mails]  $src  ->  $destPath/$n" -f $f.Items.Count) "Cyan"
    if (-not $DryRun) {
        if ($newName) { $f.Name = $newName }
        $f.MoveTo($dest)
        $st.moved++
    }
}

function Merge-Folder($src, $destPath) {
    $f = $null; try { $f = $inbox.Folders.Item($src) } catch {}
    if ($null -eq $f) { Log ("    [SKIP]    $src nicht gefunden") "DarkYellow"; $st.skipped++; return }
    $dest = Get-Folder $destPath
    if ($null -eq $dest) { Log ("    [FEHLER]  Ziel nicht gefunden: $destPath") "Red"; return }
    $cnt = $f.Items.Count
    Log ("    [MERGE {0,4} Mails]  $src  ->  $destPath" -f $cnt) "Magenta"
    if (-not $DryRun) {
        $items = @($f.Items)
        foreach ($item in $items) { try { $item.Move($dest) | Out-Null } catch {} }
        try { $f.Delete() } catch {}
        $st.merged += $cnt
    }
}

function Del-Empty($name) {
    $f = $null; try { $f = $inbox.Folders.Item($name) } catch {}
    if ($null -eq $f) { return }
    if ($f.Items.Count -gt 0) { Log ("    [SKIP-DEL] $name hat noch $($f.Items.Count) Mails") "DarkYellow"; return }
    Log ("    [LOESCHEN] $name") "DarkGray"
    if (-not $DryRun) { try { $f.Delete(); $st.deleted++ } catch {} }
}

function Classify($pattern, $destPath, $label) {
    $dest = Get-Folder $destPath
    if ($null -eq $dest) { Log ("    [FEHLER]  Ziel nicht gefunden: $destPath") "Red"; return }
    $q = [char]34; $sq = [char]39
    $filter = "@SQL=${q}urn:schemas:httpmail:fromemail${q} LIKE ${sq}%${pattern}%${sq} OR ${q}urn:schemas:httpmail:fromname${q} LIKE ${sq}%${pattern}%${sq}"
    try {
        $items = @($inbox.Items.Restrict($filter))
        if ($items.Count -eq 0) { return }
        Log ("    [{0,5} Mails]  $label  ->  $destPath" -f $items.Count) "Cyan"
        if (-not $DryRun) {
            foreach ($item in $items) { try { $item.Move($dest) | Out-Null; $st.classified++ } catch {} }
        } else { $st.classified += $items.Count }
    } catch { Log "    [FEHLER] Classify ${pattern}: $_" "Red" }
}

$col = if ($DryRun) { "Yellow" } else { "Red" }
$lbl = if ($DryRun) { "DRY RUN" } else { "LIVE - wird ausgefuehrt!" }
Write-Host ""
Write-Host "=== OUTLOOK REORGANISATION  [$lbl] ===" -ForegroundColor $col

# === PHASE 1: Hauptordner ===
Write-Host "`n[1/5] Hauptordner anlegen..." -ForegroundColor Yellow
$topFolders = @(
    "Finanzen",
    "Shopping",
    "Immobilien",
    "Kinder und Schule",
    "Beruf",
    "Vereine",
    "Ehrenamt",
    "Hobby",
    "Pers${oe}nlich",
    "Haus und Energie",
    "Digital",
    "Online-Dienste",
    "_Archiv"
)
foreach ($f in $topFolders) { Ensure-Folder $f | Out-Null }

# === PHASE 2: Ordner verschieben ===
Write-Host "`n[2/5] Ordner verschieben..." -ForegroundColor Yellow

Write-Host "  Finanzen:"
Move-Folder "PayPal Klarna"     "Finanzen"  "Zahlungen"
Move-Folder "Kreditkarte"       "Finanzen"
Move-Folder "Versicherung/Bank" "Finanzen"  "Bank und Versicherung"
Move-Folder "Rechnung"          "Finanzen"  "Rechnungen"

Write-Host "  Shopping:"
Move-Folder "Amazon"                "Shopping"
Move-Folder "Paket / Post / Retour" "Shopping"  "Pakete und Lieferungen"
Move-Folder "Einkauf"               "Shopping"  "Shops und Einkauf"
Move-Folder "^Kleinanzeigen"        "Shopping"  "Kleinanzeigen"

Write-Host "  Auto:"
Move-Folder "AutohausMolnar"  "Auto"  "Autohaus Molnar"
Move-Folder "Autoverkauf"     "Auto"  "Autoverkauf und Suche"

Write-Host "  Immobilien:"
Move-Folder "B${oe}hringen_Lichtensteinstr"  "Immobilien"  "B${oe}hringen Lichtensteinstr"
Move-Folder "Wittlingen_Schwalbenstr"         "Immobilien"  "Wittlingen Schwalbenstr"

Write-Host "  Kinder und Schule:"
Move-Folder "GEG"  "Kinder und Schule"

Write-Host "  Ehrenamt:"
Move-Folder "Ortschaftsrat"  "Ehrenamt"

Write-Host "  Hobby:"
Move-Folder "GeoCache"  "Hobby"

Write-Host "  Beruf:"
Move-Folder "Firthbauer"    "Beruf"  "Fierthbauer"
Move-Folder "Holzher"       "Beruf"
Move-Folder "Weiterbildung" "Beruf"
Move-Folder "Bewerbung"     "Beruf"

Write-Host "  Vereine:"
Move-Folder "TSV Wittlingen"            "Vereine"
Move-Folder "TSV B${oe}hringen"         "Vereine"

Write-Host "  Digital:"
Move-Folder "Lizenzen"                  "Digital"  "Software und Lizenzen"
Move-Folder "Datensicherung / Fritzbox" "Digital"  "Netzwerk und Backup"
Move-Folder "Scanned"                   "Digital"

Write-Host "  Persoenlich:"
Move-Folder "Gesundheit"  "Pers${oe}nlich"
Move-Folder "Urlaub"      "Pers${oe}nlich"  "Urlaub und Reise"
Move-Folder "Passes"      "Pers${oe}nlich"
Move-Folder "Steuer"      "Pers${oe}nlich"

Write-Host "  Online-Dienste:"
Move-Folder "Homepages"        "Online-Dienste"  "Web und Hosting"
Move-Folder "Check24"          "Online-Dienste"  "Vergleichsportale"
Move-Folder "Telekom Anbieter" "Online-Dienste"  "Telekom und Anbieter"
Move-Folder "Web.de"           "Online-Dienste"

Write-Host "  Newsletter:"
Move-Folder "Heise CT"  "Newsletter"
Move-Folder "Medium "   "Newsletter"

Write-Host "  Archiv:"
Move-Folder "GS Zainingen"  "_Archiv"

# === PHASE 3: Ordner zusammenfuehren ===
Write-Host "`n[3/5] Ordner zusammenfuehren..." -ForegroundColor Yellow

Write-Host "  Finanzen:"
Merge-Folder "_${Ue}berweisung"  "Finanzen/Zahlungen"
Merge-Folder "Bank"              "Finanzen/Bank und Versicherung"
Merge-Folder "VolkswagenBank"    "Finanzen/Bank und Versicherung"

Write-Host "  Haus und Energie:"
Merge-Folder "Haus"    "Haus und Energie"
Merge-Folder "Pellet"  "Haus und Energie"
Merge-Folder "Strom"   "Haus und Energie"

Write-Host "  Shopping:"
Merge-Folder "Einkaufsl${ae}den"  "Shopping/Shops und Einkauf"
Merge-Folder "Kundenservice"      "Shopping/Pakete und Lieferungen"
Merge-Folder "Verkauf"            "Shopping/Kleinanzeigen"

Write-Host "  Digital:"
Merge-Folder "Registrierung"  "Digital/Software und Lizenzen"
Merge-Folder "Developer"      "Digital/Software und Lizenzen"
Merge-Folder "Fax"            "Digital/Scanned"

Write-Host "  Online-Dienste:"
Merge-Folder "Strato"  "Online-Dienste/Web und Hosting"

Write-Host "  Persoenlich:"
Merge-Folder "Privat"  "Pers${oe}nlich"

# === PHASE 4: Leere Ordner loeschen ===
Write-Host "`n[4/5] Leere Ordner loeschen..." -ForegroundColor Yellow
foreach ($n in @("Jungs","BW","Support Anfrage","Tonne","Pinterest","Post","Facebook","Apps")) {
    Del-Empty $n
}

# === PHASE 5: Posteingang klassifizieren ===
Write-Host "`n[5/5] Posteingang-Mails einordnen..." -ForegroundColor Yellow

Write-Host "  Pakete:"
Classify "dhl"         "Shopping/Pakete und Lieferungen"  "DHL"
Classify "dpd"         "Shopping/Pakete und Lieferungen"  "DPD"
Classify "hermesworld" "Shopping/Pakete und Lieferungen"  "Hermes"

Write-Host "  Finanzen:"
Classify "verti.de"               "Finanzen/Bank und Versicherung"  "Verti"
Classify "barmenia"               "Finanzen/Bank und Versicherung"  "Barmenia"
Classify "sparkassen-direkt"      "Finanzen/Bank und Versicherung"  "Sparkassen-Direkt"
Classify "ksk-reutlingen"         "Finanzen/Bank und Versicherung"  "KSK Reutlingen"
Classify "sparkassenversicherung" "Finanzen/Bank und Versicherung"  "Sparkassenvers."

Write-Host "  Haus und Energie:"
Classify "netze-bw"  "Haus und Energie"  "Netze BW"
Classify "derago"    "Haus und Energie"  "Ablesung"
Classify "suewag"    "Haus und Energie"  "Suewag"

Write-Host "  Auto:"
Classify "autohaus-molnar"  "Auto/Autohaus Molnar"        "Autohaus Molnar"
Classify "automolnar"       "Auto/Autohaus Molnar"        "Autohaus Molnar"
Classify "autoscout"        "Auto/Autoverkauf und Suche"  "AutoScout24"
Classify "opel.com"         "Auto/Autoverkauf und Suche"  "Opel"
Classify "vwgroup"          "Auto/Autoverkauf und Suche"  "VW"

Write-Host "  Beruf:"
Classify "fierthbauer"  "Beruf/Fierthbauer"  "Fierthbauer"
Classify "holzher.com"  "Beruf/Holzher"      "Holzher"

Write-Host "  Persoenlich:"
Classify "abindenurlaub"  "Pers${oe}nlich/Urlaub und Reise"  "Ab in den Urlaub"

Write-Host "  Digital:"
Classify "ionos"    "Online-Dienste/Web und Hosting"  "IONOS"
Classify "synology" "Digital/Netzwerk und Backup"     "Synology"
Classify "docker"   "Digital/Software und Lizenzen"   "Docker"

# === ZUSAMMENFASSUNG ===
Write-Host ""
Write-Host "=== ZUSAMMENFASSUNG ===" -ForegroundColor $col
Write-Host ("  Neue Ordner           : {0}" -f $st.created)
Write-Host ("  Ordner verschoben     : {0}" -f $st.moved)
Write-Host ("  Mails zusammengefuehrt: {0}" -f $st.merged)
Write-Host ("  Mails klassifiziert   : {0}" -f $st.classified)
Write-Host ("  Ordner geloescht      : {0}" -f $st.deleted)
Write-Host ("  Uebersprungen         : {0}" -f $st.skipped)
if ($DryRun) {
    Write-Host ""
    Write-Host "  -> Zum Ausfuehren: .\outlook_reorganize.ps1 -Live" -ForegroundColor Yellow
} else {
    Write-Host "  Fertig!" -ForegroundColor Green
}
