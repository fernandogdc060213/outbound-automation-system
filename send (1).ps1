$ErrorActionPreference = "Stop"
Import-Module ImportExcel

# =========================
# SIGNATURE (ZERO-DISPLAY)
# =========================
function Get-DefaultSignatureHtml {
    $sigDir = Join-Path $env:APPDATA "Microsoft\Signatures"
    if (-not (Test-Path $sigDir)) { return "" }

    $htm = Get-ChildItem -Path $sigDir -Filter "*.htm" -File -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending |
           Select-Object -First 1

    if (-not $htm) { return "" }

    try {
        $html = Get-Content -Path $htm.FullName -Raw -Encoding Default
        return $html
    } catch {
        try { return (Get-Content -Path $htm.FullName -Raw) } catch { return "" }
    }
}

$SignatureHtml = Get-DefaultSignatureHtml
if ([string]::IsNullOrWhiteSpace($SignatureHtml)) {
    Write-Host "WARNING: Could not load HTML signature from %APPDATA%\Microsoft\Signatures"
}

# =========================
# CONFIG
# =========================
$ExcelPath  = Join-Path $env:USERPROFILE "Documents\53Cargo_System\data\53_cargo_Leads_DATA.xlsx"
$SheetName  = "Email_Automation"

$startHour = 6
$endHour   = 19
$spacingMinutes = 30
$sendNowGraceSeconds = 60

# =========================
# HELPERS
# =========================
function Parse-DateOrNull($v) {
    if ($null -eq $v) { return $null }
    $s = "$v".Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try { return [datetime]::Parse($s) } catch { return $null }
}

function Get-CityKey($locationRaw) {
    $s = "$locationRaw".Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($s)) { return "unknown" }
    $city = $s.Split(",")[0].Trim()
    $city = [regex]::Replace($city, "\s+", " ")
    if ([string]::IsNullOrWhiteSpace($city)) { return "unknown" }
    return $city
}

function Adjust-ToWindow([DateTime]$dt) {
    $start = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour $startHour -Minute 0 -Second 0
    $end   = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour $endHour   -Minute 0 -Second 0
    if ($dt -lt $start) { return $start }
    if ($dt -gt $end)   { return $start.AddDays(1) }
    return $dt
}

function Append-History($history, $subject) {
    $h = ""
    if ($null -ne $history) { $h = "$history".Trim() }
    $s = ""
    if ($null -ne $subject) { $s = "$subject".Trim() }

    if ([string]::IsNullOrWhiteSpace($s)) { return $h }
    if ([string]::IsNullOrWhiteSpace($h)) { return $s }
    if ($h.ToLower().Contains($s.ToLower())) { return $h }

    return "$h | $s"
}

function Get-TemplateMap($path) {
    $rows = Import-Excel -Path $path -WorksheetName "Templates"
    $map = @{}
    foreach ($r in $rows) {
        $k = "$($r.Template)".Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $map[$k] = @{
            Subject = "$($r.Subject)"
            Body    = "$($r.Body)"
        }
    }
    return $map
}

# =========================
# TEMPLATE ADVANCEMENT
# =========================
$advance = @{
    "land_intro_en"       = "land_followup_en"
    "land_followup_en"    = "land_followup2_en"
    "land_followup2_en"   = $null

    "land_intro_es"       = "land_followup_es"
    "land_followup_es"    = "land_followup2_es"
    "land_followup2_es"   = $null

    "ocean_intro_en"      = "ocean_followup_en"
    "ocean_followup_en"   = "ocean_followup2_en"
    "ocean_followup2_en"  = $null

    "ocean_intro_es"      = "ocean_followup_es"
    "ocean_followup_es"   = "ocean_followup2_es"
    "ocean_followup2_es"  = $null

    "land_intro2_en"      = "land_followup2_la_en"
    "land_followup_la_en" = $null

    "weekend_send_es"     = $null
    "weekend_send_en"     = $null

    "monday_send_es"      = $null
    "monday_send_en"      = $null
}

# =========================
# LOAD DATA
# =========================
$rows = Import-Excel -Path $ExcelPath -WorksheetName $SheetName
$tmpl = Get-TemplateMap $ExcelPath

Write-Host "Loaded rows:" $rows.Count

# =========================
# OUTLOOK
# =========================
$outlook = New-Object -ComObject Outlook.Application
$outlook.Session.Logon()
Start-Sleep -Seconds 3

$session = $outlook.Session
$account = $session.Accounts.Item(1)

$today = (Get-Date).Date
$nextSendByGroup = @{}
$sentCount = 0

# counters
$skipNoEmail = 0
$skipFinalStatus = 0
$skipNotNew = 0
$skipFutureNED = 0
$skipNoTemplate = 0
$skipNoAdvance = 0
$failedSend = 0

# =========================
# PROCESS ROWS
# =========================
foreach ($r in $rows) {

    $email = ""
    if ($r.PSObject.Properties.Match("Email").Count -gt 0) {
        $email = "$($r.Email)".Trim()
    }
    if ([string]::IsNullOrWhiteSpace($email)) {
        $skipNoEmail++
        continue
    }

    $status = "$($r.Status)".Trim().ToUpper()

    if ($status -eq "REPLIED" -or $status -eq "BOUNCED" -or $status -eq "DONE") {
        $skipFinalStatus++
        continue
    }

    if ($status -ne "NEW") {
        $skipNotNew++
        continue
    }

    $nedDt = Parse-DateOrNull $r.Next_Eligible_Date
    if ($nedDt -ne $null -and $nedDt.Year -ge 2000) {
        if ($nedDt.Date -gt $today) {
            $skipFutureNED++
            continue
        }
    }

    $templateKey = "$($r.Template)".Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($templateKey)) {
        $skipNoTemplate++
        continue
    }

    if (-not $tmpl.ContainsKey($templateKey)) {
        $skipNoTemplate++
        continue
    }

    if (-not $advance.ContainsKey($templateKey)) {
        $skipNoAdvance++
        continue
    }

    $first    = "$($r.First_Name)".Trim()
    $company  = "$($r.Company)".Trim()
    $location = "$($r.Location)".Trim()
    $last     = ""

    $cityKey = Get-CityKey $location
    $companyKey = "$company".Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($companyKey)) { $companyKey = "unknowncompany" }
    $groupKey = "$companyKey|$cityKey"

    if (-not $nextSendByGroup.ContainsKey($groupKey)) {
        $nextSendByGroup[$groupKey] = Adjust-ToWindow (Get-Date)
    }

    $scheduled = Adjust-ToWindow $nextSendByGroup[$groupKey]
    if ($scheduled -lt (Get-Date)) { $scheduled = Adjust-ToWindow (Get-Date) }

    $subjectOut = $tmpl[$templateKey].Subject
    $subjectOut = $subjectOut.Replace("{First_Name}", $first).Replace("{Last_Name}", $last).Replace("{company}", $company).Replace("{Location}", $location)

    $bodyOut = $tmpl[$templateKey].Body
    $bodyOut = $bodyOut.Replace("{First_Name}", $first).Replace("{Last_Name}", $last).Replace("{Company}", $company).Replace("{Location}", $location)

    $mail = $outlook.CreateItem(0)
    $mail.SendUsingAccount = $account
    $mail.To = $email
    $mail.Subject = $subjectOut

    $sigHtml = $SignatureHtml
    $bodyHtml = ($bodyOut -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;")
    $bodyHtml = $bodyHtml -replace "`r`n","<br>" -replace "`n","<br>"

    $mail.HTMLBody = "<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>$bodyHtml</div><br><br>" + $sigHtml

    $now = Get-Date
    $sendNow = ($scheduled -le $now.AddSeconds($sendNowGraceSeconds))

    if (-not $sendNow) {
        $mail.DeferredDeliveryTime = $scheduled
    }

    $mail.Save()
    Start-Sleep -Milliseconds 250

    $maxRetries = 4
    $sentOK = $false
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            $mail.Send()
            $sentOK = $true
            break
        } catch {
            Start-Sleep -Seconds 3
        }
    }

    if (-not $sentOK) {
        Write-Host "FAILED SEND: $email" -ForegroundColor Red
        $failedSend++
        continue
    }

    Start-Sleep -Seconds 5
    Start-Sleep -Milliseconds 350

    $nextSendByGroup[$groupKey] = Adjust-ToWindow ($scheduled.AddMinutes($spacingMinutes))

    $r.Last_Email_Sent   = $templateKey
    $r.Last_Subject_Sent = $subjectOut
    $r.Subject_History   = Append-History $r.Subject_History $subjectOut
    $r.Last_Send_Date    = $scheduled.ToString("yyyy-MM-dd HH:mm")

    $daysToNext = 7

    if ($templateKey -like "*_intro_*") {
        $daysToNext = 4
    }
    elseif ($templateKey -like "*_followup_*" -and $templateKey -notlike "*_followup2_*") {
        $daysToNext = 15
    }
    elseif ($templateKey -like "*_followup2_*") {
        $daysToNext = 9999
    }

    $r.Next_Eligible_Date = $scheduled.AddDays($daysToNext).ToString("yyyy-MM-dd")

    if ($sendNow) {
        $r.Notes = "SENT"
    } else {
        $r.Notes = "SCHEDULED"
    }

    $next = $advance[$templateKey]
    if ($null -eq $next) {
        $r.Status = "DONE"
    } else {
        $r.Template = $next
        $r.Status = "SENT"
    }

    Write-Host "OK -> $email | $($r.Notes) | $($scheduled.ToString('yyyy-MM-dd HH:mm'))"
    $sentCount++
}

Write-Host ""
Write-Host "Completed. Total scheduled/sent: $sentCount" -ForegroundColor Green
Write-Host "Skip no email:        $skipNoEmail"
Write-Host "Skip final status:    $skipFinalStatus"
Write-Host "Skip not NEW:         $skipNotNew"
Write-Host "Skip future NED:      $skipFutureNED"
Write-Host "Skip missing tmpl:    $skipNoTemplate"
Write-Host "Skip missing advance: $skipNoAdvance"
Write-Host "Failed send:          $failedSend"
Write-Host ""

# =========================
# SAVE BACK
# =========================
try {
    $rows | Export-Excel -Path $ExcelPath -WorksheetName $SheetName -AutoSize -ClearSheet
    Write-Host "Excel updated OK"
}
catch {
    Write-Host "EXPORT FAILED: $($_.Exception.Message)"
}

Read-Host "Done - press Enter"