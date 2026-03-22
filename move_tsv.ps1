$ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6)
$oe = [char]0x00F6

$dest = $inbox.Folders.Item("Vereine").Folders.Item("TSV B" + $oe + "hringen")

$q = [char]34; $sq = [char]39
$filter = "@SQL=" + $q + "urn:schemas:httpmail:fromemail" + $q + " LIKE " + $sq + "%tsv-boehringen%" + $sq
$items = @($inbox.Items.Restrict($filter))
if ($items.Count -eq 0) { Write-Host "0 Mails gefunden"; exit 0 }
foreach ($item in $items) { try { $item.Move($dest) | Out-Null } catch {} }
Write-Host ("{0} Mails  ->  Vereine/TSV Boehringen" -f $items.Count) -ForegroundColor Green
