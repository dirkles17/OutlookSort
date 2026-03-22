$ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6)
$oe = [char]0x00F6
$dest = $inbox.Folders.Item("Pers" + $oe + "nlich").Folders.Item("Urlaub und Reise")

function Move-By($pattern, $label) {
    $q = [char]34; $sq = [char]39
    $filter = "@SQL=" + $q + "urn:schemas:httpmail:fromemail" + $q + " LIKE " + $sq + "%" + $pattern + "%" + $sq
    $items = @($inbox.Items.Restrict($filter))
    if ($items.Count -eq 0) { return }
    foreach ($item in $items) { try { $item.Move($dest) | Out-Null } catch {} }
    Write-Host ("{0} Mails  ->  Persoenlich/Urlaub und Reise  [$label]" -f $items.Count) -ForegroundColor Green
}

Move-By "fewo-direkt" "Fewo-Direkt"
