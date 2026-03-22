# Outlook Deep Analyzer - Ordnerstruktur + Klassifizierung
$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox     = $namespace.GetDefaultFolder(6)

Write-Host "`n=== OUTLOOK DEEP ANALYZER ===" -ForegroundColor Cyan

# ── 1. Ordnerstruktur einlesen ──────────────────────────────────────────────
Write-Host "`n[1/3] Lese Ordnerstruktur..." -ForegroundColor Yellow

$folderStats = [ordered]@{}

function Read-Folder {
    param($folder, $depth = 0)
    $indent = "  " * $depth
    $name   = $folder.Name
    $count  = $folder.Items.Count
    $key    = ("  " * $depth) + $name
    $folderStats[$key] = @{ Name=$name; Depth=$depth; Count=$count; Folder=$folder }
    Write-Host ("{0}{1,-40} ({2} Mails)" -f $indent, $name, $count)
    foreach ($sub in $folder.Folders) { Read-Folder $sub ($depth+1) }
}

Read-Folder $inbox

# ── 2. Alle Mails samplen (max 30 pro Ordner, Posteingang max 200) ──────────
Write-Host "`n[2/3] Analysiere E-Mails..." -ForegroundColor Yellow

$allMails = [System.Collections.Generic.List[PSObject]]::new()

function Sample-Folder {
    param($folder, $folderName, $maxItems)
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    $n = [Math]::Min($maxItems, $items.Count)
    $count = 0
    foreach ($mail in $items) {
        if ($count -ge $n) { break }
        if ($mail.Class -ne 43) { $count++; continue }
        $script:allMails.Add([PSCustomObject]@{
            Folder  = $folderName
            Date    = $mail.ReceivedTime
            Sender  = $mail.SenderEmailAddress
            Name    = $mail.SenderName
            Subject = $mail.Subject
        })
        $count++
    }
}

foreach ($entry in $folderStats.GetEnumerator()) {
    $f   = $entry.Value
    $max = if ($f.Depth -eq 0) { 200 } else { 50 }
    Sample-Folder $f.Folder $f.Name $max
}

Write-Host "  $($allMails.Count) E-Mails analysiert"

# ── 3. Absender-Domains auswerten ───────────────────────────────────────────
Write-Host "`n[3/3] Werte Absender aus..." -ForegroundColor Yellow

$domainMap   = @{}
$senderMap   = @{}
$folderUsage = @{}

foreach ($m in $allMails) {
    # Domain extrahieren
    $domain = if ($m.Sender -match "@(.+)") { $matches[1].ToLower() } else { "unbekannt" }
    if (-not $domainMap[$domain])  { $domainMap[$domain]  = 0 }
    $domainMap[$domain]++

    # Absendername
    $sname = $m.Name
    if (-not $senderMap[$sname])   { $senderMap[$sname]   = 0 }
    $senderMap[$sname]++

    # Ordnernutzung
    if (-not $folderUsage[$m.Folder]) { $folderUsage[$m.Folder] = 0 }
    $folderUsage[$m.Folder]++
}

# ── AUSGABE ─────────────────────────────────────────────────────────────────

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        ORDNERSTRUKTUR (aktuell)              ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
foreach ($entry in $folderStats.GetEnumerator()) {
    $f = $entry.Value
    $indent = "  " * $f.Depth
    Write-Host ("{0}[{1,4}] {2}" -f $indent, $f.Count, $f.Name)
}

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        TOP 30 ABSENDER-DOMAINS               ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
$domainMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 30 | ForEach-Object {
    Write-Host ("  {0,-40} {1,4}x" -f $_.Key, $_.Value)
}

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        TOP 30 ABSENDER (Namen)               ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
$senderMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 30 | ForEach-Object {
    Write-Host ("  {0,-40} {1,4}x" -f $_.Key, $_.Value)
}

Write-Host "`n╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        MAILS PRO ORDNER (Sample)             ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
$folderUsage.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    Write-Host ("  {0,-40} {1,4}x" -f $_.Key, $_.Value)
}

Write-Host "`nFertig." -ForegroundColor Green
