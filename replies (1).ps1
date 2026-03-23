# ==== AUTO CLOSE EXCEL (prevents save lock) ====
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Import-Module ImportExcel

# ========= CONFIG =========
$ExcelPath = Join-Path $env:USERPROFILE "Documents\53Cargo_System\Data\53_cargo_Leads_DATA.xlsx"
$SheetName    = "Email_Automation"
$MaxInboxScan = 20000

# ========= HELPERS =========
function Normalize-Subject([string]$subject) {
    if ([string]::IsNullOrWhiteSpace($subject)) { return "" }

    $s = $subject.Trim().ToLower()

    # ASCII-only to avoid encoding issues in .ps1 files
    while ($true) {
        $new = $s -replace "^(re|fw|rv|fwd)\s*:\s*", ""
        if ($new -eq $s) { break }
        $s = $new
    }

    $s = $s -replace "\s+"," "
    return $s
}

function Parse-ExcelDate($val) {
    if ($null -eq $val) { return $null }
    $str = $val.ToString()
    if ([string]::IsNullOrWhiteSpace($str)) { return $null }
    if ($val -is [datetime]) { return [datetime]$val }

    $s = $str.Trim()
    [double]$d = 0
    if ([double]::TryParse($s, [ref]$d)) {
        try { return [datetime]::FromOADate($d) } catch {}
    }
    try { return [datetime]::Parse($s) } catch {}
    return $null
}

function Append-Note([string]$existing, [string]$toAdd) {
    if ($existing) { $existing = $existing.Trim() } else { $existing = "" }
    if ([string]::IsNullOrWhiteSpace($existing)) { return $toAdd }
    if ($existing -like "*$toAdd*") { return $existing }
    return ($existing + " | " + $toAdd).Trim(" |")
}

# ========= LOAD EXCEL =========
$rows = Import-Excel -Path $ExcelPath -WorksheetName $SheetName
Write-Host "Loaded rows:" $rows.Count

# Build lead dict: email -> row object (only those we can mark)
$leadByEmail = @{}

foreach ($r in $rows) {
    $email = ""
    if ($r.PSObject.Properties.Match("Email").Count -gt 0) {
        $email = ($r.Email | Out-String).Trim().ToLower()
    }
    if ([string]::IsNullOrWhiteSpace($email)) { continue }

    $status = ""
    if ($r.PSObject.Properties.Match("Status").Count -gt 0) {
        $status = ($r.Status | Out-String).Trim()
    }

    # Skip rows we should never touch
    if ($status -eq "REPLIED" -or $status -eq "BOUNCED" -or $status -eq "DONE") { continue }

    # Must have Last_Send_Date to compare
    $lastSend = $null
    if ($r.PSObject.Properties.Match("Last_Send_Date").Count -gt 0) {
        $lastSend = Parse-ExcelDate $r.Last_Send_Date
    }
    if ($null -eq $lastSend) { continue }

    $leadByEmail[$email] = $r
}

Write-Host "Leads eligible for reply scan:" $leadByEmail.Count

# ========= OUTLOOK =========
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
# Access custom Replied folder
$repliedFolder = $inbox.Parent.Folders.Item("Replied")


# Combine Inbox + Replied mail
$allMail = @()
$allMail += $inbox.Items
$allMail += $repliedFolder.Items

# Sort newest first
$allMail = $allMail | Sort-Object ReceivedTime -Descending

# Index latest inbound by sender SMTP
$latestInboundBySender = @{}
$limit = [Math]::Min($MaxInboxScan, $allMail.Count)

for ($i=1; $i -le $limit; $i++) {
    $m = $allMail[$i-1]
    if (-not $m -or $m.Class -ne 43) { continue }

    $rt = $null
    try { $rt = [datetime]$m.ReceivedTime } catch { continue }

    $sender = ""
    try {
        if ($m.SenderEmailType -eq "SMTP") {
            $sender = ($m.SenderEmailAddress | Out-String).Trim().ToLower()
        } else {
            $ex = $m.Sender.GetExchangeUser()
            if ($ex -and $ex.PrimarySmtpAddress) {
                $sender = ($ex.PrimarySmtpAddress | Out-String).Trim().ToLower()
            }
        }
    } catch {}

    if ([string]::IsNullOrWhiteSpace($sender)) { continue }

    if (-not $latestInboundBySender.ContainsKey($sender)) {
        $latestInboundBySender[$sender] = $rt
    }
}

Write-Host "Inbox senders indexed:" $latestInboundBySender.Count

# ========= MARK REPLIES =========
$totalReplied = 0
$nowStamp = (Get-Date -Format "yyyy-MM-dd HH:mm")

foreach ($email in $leadByEmail.Keys) {
    $r = $leadByEmail[$email]
    $lastSend = Parse-ExcelDate $r.Last_Send_Date
    if ($null -eq $lastSend) { continue }

    if ($latestInboundBySender.ContainsKey($email)) {
        $latest = [datetime]$latestInboundBySender[$email]
        if ($latest -gt $lastSend) {
            $r.Status = "REPLIED"
            $r.Notes = "REPLIED"
            $totalReplied++
        }
    }
}

Write-Host "🚀 Completed. Total marked as REPLIED:" $totalReplied

# ========= SAVE =========
$rows | Export-Excel -Path $ExcelPath -WorksheetName $SheetName -AutoSize
Write-Host "✅ Excel updated:" $ExcelPath

