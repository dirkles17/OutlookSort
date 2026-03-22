# Absender im Posteingang gruppieren (nur Posteingang, nicht Unterordner)
try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch { Write-Host "Outlook nicht offen." -ForegroundColor Red; exit 1 }

$namespace = $outlook.GetNamespace("MAPI")
$inbox     = $namespace.GetDefaultFolder(6)

$senders = @{}

foreach ($item in $inbox.Items) {
    if ($item.Class -ne 43) { continue }
    $key = $item.SenderEmailAddress.ToLower()
    if (-not $senders[$key]) { $senders[$key] = @{ name=$item.SenderName; count=0 } }
    $senders[$key].count++
}

Write-Host "`nAbsender im Posteingang (sortiert nach Anzahl):" -ForegroundColor Yellow
$senders.GetEnumerator() | Sort-Object { $_.Value.count } -Descending | ForEach-Object {
    Write-Host ("{0,4}x  {1,-35}  {2}" -f $_.Value.count, $_.Value.name, $_.Key)
}
Write-Host "`nGesamt: $($inbox.Items.Count) Mails von $($senders.Count) Absendern"
