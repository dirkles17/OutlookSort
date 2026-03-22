# Outlook Email Reader & Classifier
# Verbindet sich mit Outlook und liest E-Mails aus dem Posteingang

# Outlook COM-Objekt erstellen
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    Write-Error "Outlook konnte nicht gestartet werden: $_"
    exit 1
}

# Posteingang öffnen
$inbox = $namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
$mails = $inbox.Items
$mails.Sort("[ReceivedTime]", $true)  # Neueste zuerst

Write-Host "`n=== Outlook Email Classifier ===" -ForegroundColor Cyan
Write-Host "Posteingang: $($mails.Count) E-Mails gefunden`n" -ForegroundColor Yellow

# Klassifizierungsregeln
function Get-EmailCategory {
    param($subject, $sender)

    $subject = $subject.ToLower()
    $sender  = $sender.ToLower()

    if ($sender -match "newsletter|noreply|no-reply|marketing|info@") {
        return "Newsletter/Marketing"
    }
    elseif ($subject -match "rechnung|invoice|zahlung|payment|bestellung|order") {
        return "Rechnung/Bestellung"
    }
    elseif ($subject -match "meeting|termin|einladung|invite|calendar") {
        return "Termine/Meetings"
    }
    elseif ($subject -match "spam|gewinn|lottery|preis|winner|unsubscribe") {
        return "Spam/Werbung"
    }
    elseif ($subject -match "github|jira|gitlab|bitbucket|jenkins|alert|build|deploy") {
        return "IT/Entwicklung"
    }
    elseif ($subject -match "support|ticket|anfrage|request|hilfe|help") {
        return "Support/Anfragen"
    }
    else {
        return "Sonstige"
    }
}

# Ergebnisse sammeln
$results = @{}
$emailList = @()
$maxMails = 50  # Nur erste 50 analysieren

$count = 0
foreach ($mail in $mails) {
    if ($count -ge $maxMails) { break }
    if ($mail.Class -ne 43) { $count++; continue }  # 43 = olMail (nur echte Mails)

    $category = Get-EmailCategory -subject $mail.Subject -sender $mail.SenderEmailAddress

    $emailList += [PSCustomObject]@{
        Nr        = $count + 1
        Datum     = $mail.ReceivedTime.ToString("dd.MM.yy HH:mm")
        Absender  = $mail.SenderName.PadRight(25).Substring(0, [Math]::Min(25, $mail.SenderName.Length))
        Betreff   = if ($mail.Subject.Length -gt 45) { $mail.Subject.Substring(0,45) + "..." } else { $mail.Subject }
        Kategorie = $category
    }

    if (-not $results.ContainsKey($category)) { $results[$category] = 0 }
    $results[$category]++

    $count++
}

# Tabelle ausgeben
$emailList | Format-Table -AutoSize

# Zusammenfassung
Write-Host "`n=== Klassifizierung (letzte $maxMails E-Mails) ===" -ForegroundColor Cyan
$results.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    $bar = "#" * $_.Value
    Write-Host ("  {0,-25} {1,3}x  {2}" -f $_.Key, $_.Value, $bar) -ForegroundColor Green
}

Write-Host "`nFertig." -ForegroundColor Yellow
