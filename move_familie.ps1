# Familie-Ordner anlegen und E-Mails sortieren
# Neue Ordner: Familie/Julia, Familie/Hannes, Familie/Martina und Steffen,
#              Familie/Peter Koval, Familie/Archiv, Kirche, Persönlich/Genealogie

$ol     = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inbox  = $ol.GetNamespace("MAPI").GetDefaultFolder(6)

# ── Ordner anlegen (falls noch nicht vorhanden) ─────────────────────────────
function Ensure-Folder {
    param($parent, $name)
    $f = $null
    try { $f = $parent.Folders.Item($name) } catch {}
    if ($null -eq $f) {
        $f = $parent.Folders.Add($name)
        Write-Host "  Erstellt: $($parent.Name)/$name" -ForegroundColor Green
    }
    return $f
}

Write-Host "`n[1/3] Ordner anlegen..." -ForegroundColor Yellow

$fam = Ensure-Folder $inbox "Familie"
Ensure-Folder $fam "Julia"                  | Out-Null
Ensure-Folder $fam "Hannes"                 | Out-Null
Ensure-Folder $fam "Martina und Steffen"    | Out-Null
Ensure-Folder $fam "Peter Koval"            | Out-Null
Ensure-Folder $fam "Archiv"                 | Out-Null

Ensure-Folder $inbox "Kirche"               | Out-Null

$pers = $null
try { $pers = $inbox.Folders.Item("Persönlich") } catch {}
if ($null -ne $pers) { Ensure-Folder $pers "Genealogie" | Out-Null }

# ── Mails verschieben ────────────────────────────────────────────────────────
Write-Host "`n[2/3] E-Mails verschieben..." -ForegroundColor Yellow

function Move-BySender {
    param($senderPattern, $destPath)

    $cur = $inbox
    foreach ($part in ($destPath -split "/")) {
        $f = $null; try { $f = $cur.Folders.Item($part) } catch {}
        if ($null -eq $f) { Write-Host "  FEHLT: $destPath" -ForegroundColor Red; return }
        $cur = $f
    }
    $destFolder = $cur

    $q  = [char]34; $sq = [char]39
    $filter = "@SQL=" + $q + "urn:schemas:httpmail:fromemail" + $q +
              " LIKE " + $sq + "%" + $senderPattern + "%" + $sq
    $items = @($inbox.Items.Restrict($filter))
    if ($items.Count -eq 0) { Write-Host ("  {0,-35} 0 Mails" -f $senderPattern) -ForegroundColor DarkGray; return }
    foreach ($item in $items) { try { $item.Move($destFolder) | Out-Null } catch {} }
    Write-Host ("  {0,-35} {1,3} Mail(s)  ->  $destPath" -f $senderPattern, $items.Count) -ForegroundColor Cyan
}

# Julia (Ehefrau)
Move-BySender "schillerjulia@gmx.de"             "Familie/Julia"

# Hannes (Sohn)
Move-BySender "schiller-hannes@web.de"            "Familie/Hannes"
Move-BySender "hannes.schiller.17@gmail.com"      "Familie/Hannes"
Move-BySender "hannes@schiller-roemerstein.de"    "Familie/Hannes"

# Martina & Steffen Schmid (Schwester + Schwager)
Move-BySender "tini.schmid@icloud.com"            "Familie/Martina und Steffen"
Move-BySender "tini.schmid@me.com"                "Familie/Martina und Steffen"
Move-BySender "home@schmid-web.eu"                "Familie/Martina und Steffen"

# Peter Koval (Vetter USA)
Move-BySender "peter@peterkoval.com"              "Familie/Peter Koval"

# Archiv: Verstorbene
Move-BySender "schiller.inge@web.de"              "Familie/Archiv"

# Kirche
Move-BySender "noreply@churchtools.de"            "Kirche"

# Genealogie / Ahnenforschung
Move-BySender "reply@e.familysearch.org"          "Persönlich/Genealogie"

# Volksbank (noch nicht in Finanzen/Bank)
Move-BySender "info@voba-ermstal-alb.de"          "Finanzen/Bank und Versicherung"
Move-BySender "joshua.kehl@voba-ermstal-alb.de"   "Finanzen/Bank und Versicherung"
Move-BySender "service@union-investment.de"       "Finanzen/Bank und Versicherung"
Move-BySender "direktonlineteam@vhv.de"           "Finanzen/Bank und Versicherung"

Write-Host "`n[3/3] Fertig." -ForegroundColor Green
Write-Host "Neue Ordner: Familie (Julia/Hannes/Martina+Steffen/Peter Koval/Archiv), Kirche, Persoenlich/Genealogie`n"
