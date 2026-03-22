$ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6)
$oe = [char]0x00F6
$name = "Pers" + $oe + "nlich"
$f = $inbox.Folders.Item($name)
Write-Host ("Gefunden: " + $f.Name)
foreach ($sub in $f.Folders) { Write-Host ("  " + $sub.Name) }
