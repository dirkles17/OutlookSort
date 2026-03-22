param([string]$Pattern, [string]$Dest, [string]$Label)
$oe=[char]0x00F6; $ue=[char]0x00FC; $ae=[char]0x00E4
$Dest = $Dest -replace "oe`$","$oe" -replace "oe/","$oe/" -replace "oe ",($oe+" ")

$ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6)

$cur = $inbox
foreach ($p in ($Dest -split "/")) {
    $f = $null; try { $f = $cur.Folders.Item($p) } catch {}
    if ($null -eq $f) { Write-Host "FEHLT: $Dest" -ForegroundColor Red; exit 1 }
    $cur = $f
}
$dest = $cur

$q = [char]34; $sq = [char]39
$filter = "@SQL=" + $q + "urn:schemas:httpmail:fromemail" + $q + " LIKE " + $sq + "%" + $Pattern + "%" + $sq + " OR " + $q + "urn:schemas:httpmail:fromname" + $q + " LIKE " + $sq + "%" + $Pattern + "%" + $sq
$items = @($inbox.Items.Restrict($filter))
if ($items.Count -eq 0) { Write-Host "0 Mails: $Label" -ForegroundColor DarkYellow; exit 0 }
foreach ($item in $items) { try { $item.Move($dest) | Out-Null } catch {} }
Write-Host ("{0,3} Mails  ->  $Dest  [$Label]" -f $items.Count) -ForegroundColor Green
