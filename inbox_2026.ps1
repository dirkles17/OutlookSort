$ol    = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6)

# Nur 2026er Mails via Filter
$filter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '2026-01-01 00:00:00'"
$items2026 = $inbox.Items.Restrict($filter)

$groups = @{}
foreach ($item in $items2026) {
    if ($item.Class -ne 43) { continue }
    try {
        $key = $item.SenderEmailAddress.ToLower()
        if (-not $groups[$key]) {
            $groups[$key] = @{ name=$item.SenderName; mails=[System.Collections.Generic.List[string]]::new(); latest=$item.ReceivedTime }
        }
        $groups[$key].mails.Add(("{0}  {1}" -f $item.ReceivedTime.ToString("dd.MM"), $item.Subject))
        if ($item.ReceivedTime -gt $groups[$key].latest) { $groups[$key].latest = $item.ReceivedTime }
    } catch {}
}

$groups.GetEnumerator() | Sort-Object { $_.Value.latest } -Descending | ForEach-Object {
    $v = $_.Value
    $addr = $_.Key
    Write-Host ("`n{0,2}x  {1}  <{2}>" -f $v.mails.Count, $v.name, $addr) -ForegroundColor Yellow
    foreach ($s in $v.mails) { Write-Host "      $s" }
}
Write-Host "`nGesamt: $($groups.Count) Absender" -ForegroundColor Cyan
