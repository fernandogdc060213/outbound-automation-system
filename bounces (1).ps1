Import-Module ImportExcel

# =========================
# PATHS (robust)
# =========================
$SystemRoot = Split-Path -Parent $PSScriptRoot          # ...\53Cargo_System
$ExcelPath  = Join-Path $SystemRoot "Data\53_cargo_Leads_DATA.xlsx"

$SheetName   = "Email_Automation"
$MaxToScan   = 30000

# Scan ONLY these. We move bounces to "Bounced" so we don't re-scan them.
$FolderNames = @("Inbox","Junk Email","Deleted Items","Sent Items")

Write-Host "ExcelPath:" $ExcelPath

# =========================
# HELPERS
# =========================
function Normalize-EmailKey($val) {
    if ($null -eq $val) { return $null }
    $s = ($val | Out-String)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }

    $s = $s.ToLower()
    $s = $s -replace "[\u00A0\s]+",""   # remove NBSP + all whitespace
    $s = $s.Trim()
    $s = $s.TrimStart("(","[","{","<")
    $s = $s.TrimEnd(")","]","}",">",".",",",";",":","'","""")

    if ($s -notmatch "^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$") { return $null }
    return $s
}

function Get-SenderSmtp($msg) {
    $sender = ""

    try {
        if ($msg.SenderEmailType -eq "SMTP") {
            $sender = ($msg.SenderEmailAddress | Out-String).Trim().ToLower()
        } else {
            $ex = $null
            try { $ex = $msg.Sender.GetExchangeUser() } catch { $ex = $null }
            if ($ex -and $ex.PrimarySmtpAddress) {
                $sender = ($ex.PrimarySmtpAddress | Out-String).Trim().ToLower()
            } else {
                $sender = ($msg.SenderEmailAddress | Out-String).Trim().ToLower()
            }
        }
    } catch {
        try { $sender = ($msg.SenderEmailAddress | Out-String).Trim().ToLower() } catch { }
    }

    # Extra fallback for ReportItem / weird Outlook objects
    if ([string]::IsNullOrWhiteSpace($sender)) {
        try { $sender = ($msg.SenderName | Out-String).Trim().ToLower() } catch { }
    }

    return $sender
}

function Looks-LikeBounce($subject, $body, $sender) {
    $s = ($subject | Out-String).ToLower().Trim()
    $b = ($body    | Out-String).ToLower()
    $f = ($sender  | Out-String).ToLower().Trim()

    # 1) Strong sender signals
    $senderSignals = @(
        "postmaster",
        "mailer-daemon",
        "mail delivery subsystem",
        "delivery subsystem",
        "delivery-status",
        "mail delivery system",
        "mail system",
        "bounce",
        "microsoft outlook"
    )

    foreach ($p in $senderSignals) {
        if ($f -like "*$p*") { return $true }
    }

    # 2) Subject signals
    $subjectSignals = @(
        "undeliverable",
        "delivery has failed",
        "delivery status notification",
        "mail delivery failed",
        "returned mail",
        "non-delivery",
        "failure notice",
        "message could not be delivered",
        "your message couldn't be delivered",
        "address not found",
        "wasn't found at",
        "delivery has failed to these recipients or groups",
        "no se pudo entregar",
        "no entregado",
        "no se pudo entregar el mensaje",
        "dirección no encontrada",
        "direccion no encontrada"
    )

    foreach ($p in $subjectSignals) {
        if ($s -like "*$p*") { return $true }
    }

    # 3) Body signals
    $bodySignals = @(
        # common DSN / NDR blocks
        "delivery has failed to these recipients",
        "delivery has failed to these recipients or groups",
        "diagnostic information for administrators",
        "remote server returned",
        "transcript of session follows",
        "the following addresses had permanent fatal errors",
        "the original message was received at",

        # postfix / daemon styles
        "this is the mail system at host",
        "i'm sorry to have to inform you that your message could not be delivered",
        "recipient address rejected",
        "user unknown in virtual alias table",
        "in reply to rcpt to command",

        # exchange / outlook styles
        "the email address you entered couldn't be found",
        "recipient not found by smtp address lookup",
        "resolver.adr.recipientnotfound",

        # classic codes / terms
        "recipientnotfound",
        "resolver.adr",
        "rejected your message",
        "user unknown",
        "mailbox unavailable",
        "unknown to address",
        "recipient address rejected",
        "address not found",

        # codes
        "550",
        "551",
        "552",
        "553",
        "554",
        "4.2.2",
        "5.0.0",
        "5.1.1",
        "5.1.10",
        "5.2.1",
        "5.4.1",
        "5.4.14",
        "5.7.1",

        # hop count / loops / policy
        "hop count exceeded",
        "mail loop",
        "possible mail loop",
        "blocked",
        "blocked by policy",
        "message refused",
        "spam",
        "security policy"
    )

    foreach ($p in $bodySignals) {
        if ($b -like "*$p*") { return $true }
    }

    return $false
}

function Extract-FailedRecipient($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $patterns = @(
        # postfix / daemon formats
        "<\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>:\s*host\s+.*?\s+said:",
        "<\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>:\s*Recipient address rejected",
        "<\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>:\s*User unknown",
        "<\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>",

        # exchange / outlook formats
        "Delivery has failed to these recipients or groups:\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "The email address you entered couldn't be found.*?([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})",
        "recipient not found by smtp address lookup.*?([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})",

        # modern templates
        "wasn't delivered to\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "message wasn't delivered to\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "your message to\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?\s*couldn't be delivered",
        "message to\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?\s*couldn't be delivered",

        # classic DSN blocks
        "rejected your message to the following email addresses:\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "The following addresses had permanent fatal errors\s*-+\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "---- The following addresses had permanent fatal errors ----\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",

        # spanish
        "No se pudo entregar.*?:\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "No entregado.*?:\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?",
        "direccion no encontrada.*?:\s*<?\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})\s*>?"
    )

    foreach ($p in $patterns) {
        $m = [regex]::Match($text, $p, "IgnoreCase, Singleline")
        if ($m.Success -and $m.Groups.Count -ge 2) {
            $e = Normalize-EmailKey $m.Groups[1].Value
            if ($e) { return $e }
        }
    }

    # fallback: first real-looking email not ours / not system addresses
    $matches = [regex]::Matches($text, "([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", "IgnoreCase")
    foreach ($m in $matches) {
        $e = Normalize-EmailKey $m.Groups[1].Value
        if (-not $e) { continue }
        if ($e -like "*@53cargo.com") { continue }
        if ($e -like "support@*" -or $e -like "noreply@*" -or $e -like "no-reply@*" -or $e -like "postmaster@*" -or $e -like "mailer-daemon@*") { continue }
        if ($e -match "\.(png|jpg|jpeg|gif|svg|webp)$") { continue }
        return $e
    }

    return $null
}

# =========================
# LOAD EXCEL
# =========================
if (-not (Test-Path $ExcelPath)) {
    Write-Host "File not found:" $ExcelPath
    exit 1
}

$rows = Import-Excel -Path $ExcelPath -WorksheetName $SheetName
Write-Host "Loaded rows:" $rows.Count

$emailToIndex = @{}
for ($i = 0; $i -lt $rows.Count; $i++) {
    $ek = $null
    if ($rows[$i].PSObject.Properties.Match("Email").Count -gt 0) {
        $ek = Normalize-EmailKey $rows[$i].Email
    }
    if (-not $ek) { continue }
    if (-not $emailToIndex.ContainsKey($ek)) { $emailToIndex[$ek] = $i }
}
Write-Host "Indexed emails:" $emailToIndex.Count

# =========================
# OUTLOOK
# =========================
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")

function Get-Folder($name) {
    if ($name -eq "Inbox")         { return $ns.GetDefaultFolder(6) }
    if ($name -eq "Junk Email")    { return $ns.GetDefaultFolder(23) }
    if ($name -eq "Deleted Items") { return $ns.GetDefaultFolder(3) }
    if ($name -eq "Sent Items")    { return $ns.GetDefaultFolder(5) }
    return $null
}

function Get-OrCreateBouncedFolder() {
    # Creates Bounced at mailbox root
    $root = $ns.GetDefaultFolder(6).Parent
    try {
        return $root.Folders.Item("Bounced")
    } catch {
        return $root.Folders.Add("Bounced")
    }
}

$BouncedFolder = Get-OrCreateBouncedFolder

$marked     = 0
$seen       = 0
$notMatched = 0
$moved      = 0

foreach ($fname in $FolderNames) {
    $folder = Get-Folder $fname
    if (-not $folder) { continue }

    Write-Host "Scanning folder:" $folder.Name
    $items = $folder.Items
    try { $items.Sort("[ReceivedTime]", $true) | Out-Null } catch { }

    $limit = [Math]::Min($MaxToScan, $items.Count)

    # IMPORTANT: loop backwards when moving items out of folder
    for ($i = $limit; $i -ge 1; $i--) {
        $msg = $null
        try { $msg = $items.Item($i) } catch { continue }
        if (-not $msg) { continue }

        # Allow MailItem (43) and ReportItem (46)
        try {
            if ($msg.Class -ne 43 -and $msg.Class -ne 46) { continue }
        } catch { continue }

        $subject = ""
        $sender  = ""
        $body    = ""
        $html    = ""

        try { $subject = ($msg.Subject | Out-String).Trim() } catch { $subject = "" }
        try { $sender  = Get-SenderSmtp $msg } catch { $sender = "" }
        try { $body    = ($msg.Body | Out-String) } catch { $body = "" }

        if (-not (Looks-LikeBounce $subject $body $sender)) { continue }

        $seen++

        $failed = Extract-FailedRecipient $body

        if (-not $failed) {
            try { $html = ($msg.HTMLBody | Out-String) } catch { $html = "" }
            $failed = Extract-FailedRecipient $html
        }

        if ($failed -and $emailToIndex.ContainsKey($failed)) {
            $idx = $emailToIndex[$failed]
            $rows[$idx].Status = "BOUNCED"
            $rows[$idx].Notes  = "BOUNCED"
            $marked++
            Write-Host "Matched + marked BOUNCED:" $failed
        } else {
            $notMatched++
            if ($failed) {
                if ($notMatched -le 15) { Write-Host "Bounce detected but not matched in sheet:" $failed }
            } else {
                if ($notMatched -le 15) { Write-Host "Bounce detected but failed recipient could not be extracted:" $subject }
            }
        }

        # Move ALL detected bounce emails so they don't get re-scanned
        try {
            Write-Host "Moving bounce message to folder 'Bounced':" $subject
            $null = $msg.Move($BouncedFolder)
            $moved++
        } catch {
            Write-Host "WARNING: could not move to Bounced folder:" $($_.Exception.Message)
        }
    }
}

Write-Host "Bounce items seen:" $seen
Write-Host "Not matched in sheet / no failed-recipient:" $notMatched
Write-Host "Moved to Bounced folder:" $moved
Write-Host "Completed. Total rows marked as BOUNCED:" $marked

try {
    $rows | Export-Excel -Path $ExcelPath -WorksheetName $SheetName -AutoSize -ClearSheet
    Write-Host "Excel updated OK"
} catch {
    Write-Host "EXPORT FAILED: $($_.Exception.Message)"
}