# Top Absender im Posteingang analysieren
try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch { Write-Host "Outlook nicht offen." -ForegroundColor Red; exit 1 }

$namespace = $outlook.GetNamespace("MAPI")
$inbox     = $namespace.GetDefaultFolder(6)

$senders = @{}

function Scan-Folder($folder) {
    try {
        foreach ($item in $folder.Items) {
            if ($item.Class -ne 43) { continue }
            $key = $item.SenderEmailAddress.ToLower()
            if (-not $senders[$key]) { $senders[$key] = @{ name=$item.SenderName; count=0 } }
            $senders[$key].count++
        }
    } catch {}
    foreach ($sub in $folder.Folders) { Scan-Folder $sub }
}

Write-Host "Scanne alle Ordner..." -ForegroundColor Cyan
Scan-Folder $inbox

Write-Host "`nTop 40 Absender:" -ForegroundColor Yellow
$senders.GetEnumerator() | Sort-Object { $_.Value.count } -Descending | Select-Object -First 40 | ForEach-Object {
    Write-Host ("{0,5}x  {1,-35}  {2}" -f $_.Value.count, $_.Value.name, $_.Key)
}
