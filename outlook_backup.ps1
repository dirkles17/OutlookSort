$src = 'H:\_Outlook\Outlook_WIN11.pst'
$dst = 'H:\_Outlook\Outlook_WIN11_Backup_20260321_0851.pst'

if (-not (Test-Path $src)) {
    Write-Host "FEHLER: PST-Datei nicht gefunden: $src" -ForegroundColor Red
    exit 1
}
if (Test-Path $dst) {
    Write-Host "Backup existiert bereits: $dst" -ForegroundColor Yellow
    exit 0
}

$sizeMB = [math]::Round((Get-Item $src).Length / 1MB)
Write-Host "Quelle : $src ($sizeMB MB)" -ForegroundColor Cyan
Write-Host "Ziel   : $dst" -ForegroundColor Cyan
Write-Host "Starte Sicherung..." -ForegroundColor Yellow

Copy-Item -Path $src -Destination $dst -ErrorAction Stop

$dstSizeMB = [math]::Round((Get-Item $dst).Length / 1MB)
Write-Host "Backup erfolgreich erstellt! ($dstSizeMB MB)" -ForegroundColor Green
