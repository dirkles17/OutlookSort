# Mails eines bestimmten Absenders suchen und löschen
param(
    [string]$Sender = "pearl",
    [switch]$Live
)
$DryRun = -not $Live

# ── Outlook starten falls nicht offen, dann verbinden ──────────
function Get-OutlookInstance {
    # Erstmal versuchen die laufende Instanz zu holen
    try {
        $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Write-Host "Verbunden mit laufendem Outlook." -ForegroundColor Green
        return $ol
    } catch {}

    # Nicht offen → sauber als Prozess starten
    Write-Host "Starte Outlook..." -ForegroundColor Yellow
    $outlookExe = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue)."(default)"
    if (-not $outlookExe) { $outlookExe = "outlook.exe" }
    Start-Process $outlookExe
    Write-Host "Warte bis Outlook bereit ist..." -ForegroundColor Yellow

    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 2
        try {
            $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            Start-Sleep -Seconds 3  # kurz warten bis MAPI vollständig geladen
            Write-Host "Outlook bereit." -ForegroundColor Green
            return $ol
        } catch {}
        Write-Host "  ... ($($i*2)s)" -ForegroundColor DarkGray
    }
    Write-Host "FEHLER: Outlook konnte nicht gestartet werden." -ForegroundColor Red
    exit 1
}

$outlook   = Get-OutlookInstance
$namespace = $outlook.GetNamespace("MAPI")
$inbox     = $namespace.GetDefaultFolder(6)

$found = [System.Collections.Generic.List[object]]::new()

$label = if ($DryRun) { "DRY RUN" } else { "LIVE - wird gelöscht!" }
Write-Host "`n=== Mails löschen: '$Sender'  [$label] ===" -ForegroundColor $(if($DryRun){"Yellow"}else{"Red"})
Write-Host "Suche im Posteingang..." -ForegroundColor Cyan

try {
    $filter = "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%$Sender%' OR ""urn:schemas:httpmail:fromname"" LIKE '%$Sender%'"
    $items  = $inbox.Items.Restrict($filter)
    foreach ($item in $items) { $found.Add($item) }
} catch {}

Write-Host "$($found.Count) Mails gefunden:`n" -ForegroundColor Yellow

$found | Sort-Object ReceivedTime -Descending | ForEach-Object {
    Write-Host ("  {0}  {1,-25}  {2}" -f `
        $_.ReceivedTime.ToString("dd.MM.yy"), `
        $_.SenderName.PadRight(25).Substring(0,[Math]::Min(25,$_.SenderName.Length)), `
        $_.Subject)
}

if ($DryRun) {
    Write-Host "`n→ Zum Löschen: .\outlook_delete_sender.ps1 -Sender '$Sender' -Live" -ForegroundColor Yellow
} else {
    Write-Host "`nLösche $($found.Count) Mails..." -ForegroundColor Red
    $count = 0
    foreach ($item in $found) { try { $item.Delete(); $count++ } catch {} }
    $trash = $namespace.GetDefaultFolder(3)
    $trashItems = @($trash.Items)
    foreach ($item in $trashItems) { try { $item.Delete() } catch {} }
    Write-Host "$count Mails endgültig gelöscht." -ForegroundColor Green
}
