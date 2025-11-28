<#
Copyright (c) 2025 Bensha1113
Licensed under the MIT License. See LICENSE or README for details.

.SYNOPSIS
    Export a high-quality PDF report of all licensed Microsoft 365 users and license breakdown.

.DESCRIPTION
    - Uses Microsoft Graph PowerShell (Microsoft.Graph) to enumerate users and subscribed SKUs.
    - Produces a styled HTML report and converts it to PDF using the best available engine:
        wkhtmltopdf (recommended), Edge/Chrome headless --print-to-pdf, or falls back to saving HTML.
    - Requires an admin account with consent to read users/subscribed SKUs (Directory.Read.All / User.Read.All).

.USAGE
    Save as Export-M365LicensedUsersReport.ps1 and run from an elevated PowerShell session:
        pwsh .\Export-M365LicensedUsersReport.ps1 -OutputPdfPath "C:\Reports\M365-LicensedUsers.pdf"

.PARAMETER OutputPdfPath
    Full path to the output PDF. Default: .\M365-LicensedUsers.pdf

.NOTES
    - If Microsoft.Graph is not installed, the script will offer to install it.
    - You will be prompted to sign-in (admin consent may be required).
#>

param(
        [string]$OutputPdfPath,
        [ValidateSet("English", "Swedish", "Bilingual")]
        [string]$Language = "English",
        [switch]$UseWkHtml,
        [ValidateRange(1, 3650)]
        [int]$InactiveThresholdDays = 90,
        [string]$DecisionOverrideCsvPath
)

# Resolve report language directly from parameter (defaults to English)
$script:ReportLanguage = $Language

# Helper to render localized labels so the layout stays concise while supporting ENG/SWE
function Get-LocalizedText {
        param(
                [Parameter(Mandatory)] [string]$English,
                [Parameter(Mandatory)] [string]$Swedish
        )
        switch ($script:ReportLanguage) {
                'English' { return $English }
                'Swedish' { return $Swedish }
                default { return "$English<br><span class='subtitle'>$Swedish</span>" }
        }
}

function Get-LicenseNameMap {
        param([string]$CatalogUrl)
        $map = @{}
        if (-not $CatalogUrl) { return $map }

        $tempCatalog = Join-Path $env:TEMP "m365-license-catalog.csv"
        try {
                Invoke-WebRequest -Uri $CatalogUrl -OutFile $tempCatalog -UseBasicParsing -ErrorAction Stop | Out-Null
                $csv = Import-Csv -Path $tempCatalog -ErrorAction Stop
                foreach ($row in $csv) {
                        $stringId = $row.String_Id
                        $productName = $row.Product_Display_Name
                        if ($stringId -and $productName) {
                                $map[$stringId] = $productName
                        }
                }
        }
        catch {
                Write-Warning "Could not download friendly license names: $_"
        }
        finally {
                if (Test-Path $tempCatalog) { Remove-Item $tempCatalog -Force -ErrorAction SilentlyContinue }
        }

        return $map
}

function Install-ModuleIfMissing {
        param([string]$Name)
        if (-not (Get-Module -ListAvailable -Name $Name)) {
                Write-Host "Module $Name not found. Installing..." -ForegroundColor Yellow
                Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
        }
}

function Get-UserDocumentsFolder {
        $folders = @([System.Environment]::GetFolderPath('MyDocuments'),
                $env:OneDrive,
                $env:OneDriveConsumer)
        if ($env:USERPROFILE) {
                $folders += (Join-Path $env:USERPROFILE 'Documents')
        }
        foreach ($path in $folders) {
                if ($path -and (Test-Path $path)) { return $path }
        }
        return (Get-Location).Path
}

# Ensure Microsoft Graph module
Write-Host "Checking for Microsoft.Graph module..." -ForegroundColor Cyan
Install-ModuleIfMissing -Name Microsoft.Graph

# Import module
Write-Host "Importing Microsoft.Graph modules..." -ForegroundColor Cyan
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Write-Host "Modules imported successfully." -ForegroundColor Green

# Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$scopes = @("User.Read.All", "Directory.Read.All", "AuditLog.Read.All")
Connect-MgGraph -Scopes $scopes

# Get subscribed SKUs to map skuId -> skuPartNumber
Write-Host "Retrieving subscribed SKUs..." -ForegroundColor Cyan
$skuList = Get-MgSubscribedSku -All
$skuMap = @{}
foreach ($s in $skuList) {
        if ($s.SkuId) {
                $skuMap[$s.SkuId] = $s.SkuPartNumber
        }
}

$licenseCatalogUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$licenseNameMap = Get-LicenseNameMap -CatalogUrl $licenseCatalogUrl
if ($licenseNameMap.Count -gt 0) {
        Write-Host "Loaded $($licenseNameMap.Count) friendly license names" -ForegroundColor Cyan
}

# Get all users with assignedLicenses property
Write-Host "Retrieving all users (this may take a while)..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property "AssignedLicenses,DisplayName,UserPrincipalName,Mail" -ConsistencyLevel eventual
Write-Host "Retrieved $($users.Count) users" -ForegroundColor Cyan

$tenantDomain = $null
try {
        $orgResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType PSObject
        $orgItem = $orgResponse.value | Select-Object -First 1
        if ($orgItem -and $orgItem.verifiedDomains) {
                $defaultDomain = $orgItem.verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1
                if (-not $defaultDomain) { $defaultDomain = $orgItem.verifiedDomains | Where-Object { $_.isInitial } | Select-Object -First 1 }
                if (-not $defaultDomain) { $defaultDomain = $orgItem.verifiedDomains | Select-Object -First 1 }
                if ($defaultDomain) { $tenantDomain = $defaultDomain.name }
        }
}
catch {
        Write-Warning "Could not retrieve tenant domain information: $_"
}

if (-not $tenantDomain) {
        $firstUpn = $users | Where-Object { $_.UserPrincipalName } | Select-Object -First 1
        if ($firstUpn -and $firstUpn.UserPrincipalName -and $firstUpn.UserPrincipalName -match '@(.+)$') {
                $tenantDomain = $Matches[1]
        }
}

if (-not $tenantDomain) { $tenantDomain = "default-domain" }
$safeTenantDomain = ($tenantDomain -replace '[\\/:*?"<>|]', '_')
$documentsFolder = Get-UserDocumentsFolder
$reportRoot = Join-Path $documentsFolder 'User Reports'
if (-not (Test-Path $reportRoot)) { New-Item -Path $reportRoot -ItemType Directory -Force | Out-Null }
$domainFolder = Join-Path $reportRoot $safeTenantDomain
$reportGeneratedAt = Get-Date
$timestampFolderName = $reportGeneratedAt.ToString('yyyyMMdd-HHmmss')
$timestampFolder = Join-Path $domainFolder $timestampFolderName
if (-not (Test-Path $domainFolder)) { New-Item -Path $domainFolder -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $timestampFolder)) { New-Item -Path $timestampFolder -ItemType Directory -Force | Out-Null }
Write-Host "Report workspace: $timestampFolder" -ForegroundColor Cyan
$customOutputPathProvided = $PSBoundParameters.ContainsKey('OutputPdfPath') -and $OutputPdfPath
if (-not $customOutputPathProvided) {
        $OutputPdfPath = Join-Path $timestampFolder 'M365-LicensedUsers.pdf'
}

# Load manual decision overrides if supplied
$decisionOverrides = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($DecisionOverrideCsvPath) {
        if (Test-Path $DecisionOverrideCsvPath) {
                try {
                        $rawOverrides = Import-Csv -Path $DecisionOverrideCsvPath -ErrorAction Stop
                        foreach ($row in $rawOverrides) {
                                $id = $row.UPN
                                if (-not $id -and $row.Email) { $id = $row.Email }
                                if (-not $id -and $row.UserPrincipalName) { $id = $row.UserPrincipalName }
                                if (-not $id) { continue }
                                $rawDecision = $row.Action
                                if (-not $rawDecision -and $row.Decision) { $rawDecision = $row.Decision }
                                if (-not $rawDecision -and $row.Status) { $rawDecision = $row.Status }
                                if (-not $rawDecision) { continue }
                                switch -Regex ($rawDecision.ToString().Trim().ToLowerInvariant()) {
                                        '^(keep|retain|green|stay)$' { $decisionOverrides[$id] = 'keep'; continue }
                                        '^(delete|remove|drop|red)$' { $decisionOverrides[$id] = 'delete'; continue }
                                }
                        }
                        Write-Host "Loaded $($decisionOverrides.Count) manual decisions from $DecisionOverrideCsvPath" -ForegroundColor Cyan
                }
                catch {
                        Write-Warning "Could not load override CSV: $_"
                }
        }
        else {
                Write-Warning "Decision override file not found: $DecisionOverrideCsvPath"
        }
}

# Pull sign-in history from audit logs (standard retention ~30 days)
$signInLookup = @{}
$signInLookbackDays = [math]::Min([math]::Max($InactiveThresholdDays, 30), 365)
$logSince = (Get-Date).AddDays(-1 * $signInLookbackDays)
$filterTimestamp = $logSince.ToUniversalTime().ToString("o")
Write-Host "Retrieving sign-in logs since $filterTimestamp..." -ForegroundColor Cyan
try {
        $signInLogs = Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $filterTimestamp" -Property userPrincipalName,createdDateTime
        foreach ($entry in $signInLogs) {
                $upn = $entry.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                if (-not $entry.CreatedDateTime) { continue }
                $entryTime = [datetime]$entry.CreatedDateTime
                if (-not $signInLookup.ContainsKey($upn) -or $signInLookup[$upn] -lt $entryTime) {
                        $signInLookup[$upn] = $entryTime
                }
        }
        Write-Host "Captured sign-in timestamps for $($signInLookup.Count) users" -ForegroundColor Cyan
}
catch {
        Write-Warning "Could not retrieve sign-in logs: $_"
}

# Build objects for licensed users
$inactiveCutoff = (Get-Date).AddDays(-1 * [math]::Abs($InactiveThresholdDays))
$keepLabel = Get-LocalizedText 'Keep' 'Behåll'
$deleteLabel = Get-LocalizedText 'Delete' 'Ta bort'

$licensedUsers = $users | ForEach-Object {
        $u = $_
        $licenses = @()
        
        # AssignedLicenses can be in the main object or AdditionalProperties
        $assignedLicenses = $null
        $lastLogin = $null
        if ($u.AssignedLicenses) {
                $assignedLicenses = $u.AssignedLicenses
        } elseif ($u.AdditionalProperties -and $u.AdditionalProperties.ContainsKey("assignedLicenses")) {
                $assignedLicenses = $u.AdditionalProperties["assignedLicenses"]
        }

        $lastLogin = $null
        if ($u.UserPrincipalName -and $signInLookup.ContainsKey($u.UserPrincipalName)) {
                $lastLogin = $signInLookup[$u.UserPrincipalName]
        } elseif ($u.Mail -and $signInLookup.ContainsKey($u.Mail)) {
                $lastLogin = $signInLookup[$u.Mail]
        }

        $neverLabel = Get-LocalizedText 'Never' 'Aldrig'
        $lastLoginDisplay = if ($lastLogin) { $lastLogin.ToUniversalTime().ToString("yyyy-MM-dd HH:mm 'UTC'") } else { $neverLabel }
        $manualStatus = $null
        foreach ($identifier in @($u.UserPrincipalName, $u.Mail)) {
                if (-not $identifier) { continue }
                if ($decisionOverrides.ContainsKey($identifier)) {
                        $manualStatus = $decisionOverrides[$identifier]
                        break
                }
        }

        $status = if ($manualStatus) {
                $manualStatus
        } elseif ($lastLogin -and $lastLogin -ge $inactiveCutoff) {
                'keep'
        } else {
                'delete'
        }
        $statusLabel = if ($status -eq 'keep') { $keepLabel } else { $deleteLabel }
        
        if ($assignedLicenses) {
                foreach ($al in $assignedLicenses) {
                        # assignedLicenses has SkuId
                        $skuId = if ($al.SkuId) { $al.SkuId } else { $al.skuId }
                        if ($skuId) {
                                $skuCode = $skuMap[$skuId]
                                if (-not $skuCode) { $skuCode = $skuId }
                                $friendlyName = $licenseNameMap[$skuCode]
                                $skuName = if ($friendlyName) { $friendlyName } else { $skuCode }
                                $licenses += $skuName
                        }
                }
        }
        
        [PSCustomObject]@{
                DisplayName = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                Mail = $u.Mail
                LicenseCount = ($licenses | Select-Object -Unique).Count
                Licenses = ($licenses | Select-Object -Unique) -join ", "
                LastLogin = $lastLogin
                LastLoginDisplay = $lastLoginDisplay
                LifecycleStatus = $status
                LifecycleStatusLabel = $statusLabel
        }
} | Where-Object { $_.LicenseCount -gt 0 } | Sort-Object -Property DisplayName

Write-Host "Found $($licensedUsers.Count) licensed users" -ForegroundColor Cyan

$totalUsers = $users.Count
$totalLicensed = $licensedUsers.Count
$totalUnlicensed = $totalUsers - $totalLicensed
$keepCount = ($licensedUsers | Where-Object { $_.LifecycleStatus -eq 'keep' }).Count
$deleteCount = $totalLicensed - $keepCount

# SKU breakdown
$skuBreakdown = $licensedUsers |
        ForEach-Object { $_.Licenses -split ",\s*" } |
        Where-Object { $_ -ne "" } |
        Group-Object |
        Sort-Object Count -Descending |
        Select-Object @{n='Sku';e={$_.Name}}, @{n='Users';e={$_.Count}}

# Build HTML
$now = $reportGeneratedAt.ToString("u")
$reportHtml = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>M365 Licensed Users Report</title>
<style>
        @page { size: A4; margin: 12mm; }
        :root {
                --accent: #2563eb;
                --accent-soft: #dbeafe;
                --text-main: #111827;
                --text-muted: #6b7280;
                --border: #e5e7eb;
                --card-bg: #f8fafc;
        }
        body {
                font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                color: var(--text-main);
                background: #eef2ff;
                margin: 0;
                padding: 20px;
        }
        .report {
                max-width: 960px;
                margin: 0 auto;
                background: #fff;
                padding: 20px 22px;
                border-radius: 16px;
                box-shadow: 0 18px 38px -16px rgba(15, 23, 42, 0.25);
        }
        .hero h1 { margin: 0; font-size: 28px; color: var(--accent); }
        .meta { color: var(--text-muted); margin-top: 6px; }
        .toolbar { margin-top: 16px; }
        .print-btn {
                display: inline-flex;
                align-items: center;
                gap: 6px;
                padding: 9px 14px;
                font-size: 13px;
                border-radius: 999px;
                border: 1px solid var(--accent);
                color: var(--accent);
                background: transparent;
                cursor: pointer;
                transition: background 0.2s;
        }
        .print-btn:hover { background: var(--accent-soft); }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin: 18px 0 14px; page-break-inside: avoid; }
        .metric {
                background: var(--card-bg);
                border: 1px solid var(--border);
                border-radius: 12px;
                padding: 14px 16px;
                page-break-inside: avoid;
        }
        .metric strong { font-size: 13px; color: var(--text-muted); }
        .metric-value { font-size: 26px; font-weight: 600; margin-top: 4px; }
        .card-section { background: var(--card-bg); border: 1px solid var(--border); border-radius: 18px; padding: 22px 24px; margin-bottom: 26px; page-break-inside: avoid; }
        .card-section h2 { margin-top: 0; color: var(--accent); font-size: 19px; }
        table { width:100%; border-collapse: collapse; font-size:11px; }
        thead { display: table-header-group; }
        tfoot { display: table-footer-group; }
        th, td { padding:8px 9px; border-bottom:1px solid var(--border); text-align:left; vertical-align:top; page-break-inside: avoid; }
        th { background:#eef2ff; font-weight:600; }
        tbody tr:nth-child(even) { background:#f9fafb; }
        tbody tr.row-keep { background: rgba(16, 185, 129, 0.15) !important; }
        tbody tr.row-delete { background: rgba(248, 113, 113, 0.18) !important; }
        .table-wrapper { overflow:hidden; border-radius: 14px; border:1px solid var(--border); background:#fff; }
        .note { font-size:12px; color: var(--text-muted); margin-bottom: 18px; }
        .badge {
                display:inline-flex;
                align-items:center;
                gap:6px;
                padding:3px 9px;
                border-radius:999px;
                font-size:11px;
                font-weight:600;
        }
        .badge.keep { background:#dcfce7; color:#166534; }
        .badge.delete { background:#fee2e2; color:#991b1b; }
        footer { margin-top: 32px; color: var(--text-muted); font-size:11px; text-align:center; }
        .small { font-size:11px; color: var(--text-muted); }
        .subtitle { font-size:11px; color: var(--text-muted); display:block; }
        body.legacy-mode {
                background:#fff;
                font-family:"Segoe UI", Tahoma, Arial, sans-serif;
                color:#222;
        }
        body.legacy-mode .report {
                border-radius: 0;
                box-shadow: none;
                padding: 12mm;
                max-width: none;
        }
        body.legacy-mode .hero h1 {
                font-size: 22px;
                color:#222;
                margin-bottom: 4px;
        }
        body.legacy-mode .meta {
                color:#555;
                margin-bottom: 20px;
        }
        body.legacy-mode .summary {
                display:flex;
                gap:12px;
                flex-wrap:wrap;
                margin-bottom: 12px;
        }
        body.legacy-mode .metric {
                background:#f7f9fb;
                border-radius:6px;
                border:none;
                box-shadow:0 1px 0 rgba(0,0,0,0.05);
                padding:8px 12px;
        }
        body.legacy-mode .metric strong { font-size:12px; color:#555; }
        body.legacy-mode .metric-value { font-size:18px; margin-top:2px; }
        body.legacy-mode .card-section {
                border:none;
                background:transparent;
                padding:0;
                margin-bottom:20px;
        }
        body.legacy-mode .note {
                font-size:11px;
                color:#555;
                margin-bottom:18px;
        }
        body.legacy-mode .card-section h2 {
                color:#222;
                font-size:18px;
        }
        body.legacy-mode .table-wrapper {
                border:none;
                border-radius:0;
                box-shadow:none;
        }
        body.legacy-mode table { font-size:10.5px; }
        body.legacy-mode th, body.legacy-mode td {
                padding:6px 8px;
                border-bottom:1px solid #e6e9ee;
                text-align:left;
        }
        body.legacy-mode th {
                background:#f0f4f8;
                color:#222;
        }
        body.legacy-mode tbody tr:nth-child(even) { background:#fff; }
        body.legacy-mode tbody tr.row-keep { background:#ebfbee !important; }
        body.legacy-mode tbody tr.row-delete { background:#fff1f2 !important; }
        body.legacy-mode .badge.keep { background:#b7f7c3; color:#14532d; }
        body.legacy-mode .badge.delete { background:#ffc9c9; color:#7f1d1d; }
        body.legacy-mode .badge { font-size:11px; }
        body.legacy-mode footer {
                margin-top:20px;
                color:#666;
                font-size:11px;
                text-align:left;
        }
        body.legacy-mode .print-btn { display:none; }
        @media print {
                body { background:#fff; padding:0; font-size:10.5px; }
                .report { box-shadow:none; border-radius:0; padding:12mm; }
                table { font-size:9.5px; }
                th, td { padding:5px 7px; }
                .card-section, .metric { background:#fff; }
                .card-section { margin-bottom: 12px; }
                .summary { margin-bottom: 8px; }
                h2 { page-break-after: avoid; }
                .toolbar { display:none; }
        }
        tr { page-break-inside: avoid; }
        .card-section { page-break-inside: avoid; }
</style>
</head>
<body>
<div class="report">
<header class="hero">
        <h1>$(Get-LocalizedText 'Microsoft 365 Licensed Users Report' 'Rapport över licensierade Microsoft 365-användare')</h1>
        <div class="meta">$(Get-LocalizedText 'Generated' 'Genererad'): $now</div>
        <div class="toolbar"><button id="legacyPrintButton" class="print-btn" type="button">🖨️ $(Get-LocalizedText 'Print classic layout' 'Skriv ut klassiskt läge')</button></div>
</header>

<section class="summary">
        <div class="metric"><strong>$(Get-LocalizedText 'Total users' 'Totalt antal användare')</strong><div class="metric-value">$totalUsers</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Licensed users' 'Licensierade användare')</strong><div class="metric-value">$totalLicensed</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Unlicensed users' 'Olicensierade användare')</strong><div class="metric-value">$totalUnlicensed</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Recommended keep' 'Föreslås behållas')</strong><div class="metric-value" id="metric-keep">$keepCount</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Recommended delete' 'Föreslås tas bort')</strong><div class="metric-value" id="metric-delete">$deleteCount</div></div>
</section>

<div class="note">$(Get-LocalizedText "Keep = last login within $InactiveThresholdDays days (based on available sign-in logs). Delete = older activity or never signed in. Manual overrides from CSV take priority. You can also click the action badge below to adjust the visual recommendation before printing." "Behåll = senaste inloggning inom $InactiveThresholdDays dagar (baserat på tillgängliga inloggningsloggar). Ta bort = äldre aktivitet eller aldrig inloggad. Manuella CSV-överstyrningar har företräde. Du kan också klicka på åtgärdsbadgen för att justera rekommendationen innan utskrift.")</div>

<section class="card-section">
        <h2>$(Get-LocalizedText 'License SKU breakdown' 'Licensöversikt per SKU')</h2>
    <div class="table-wrapper">
        <table class="sku-table">
                <thead><tr><th>$(Get-LocalizedText 'SKU' 'SKU')</th><th>$(Get-LocalizedText 'User count' 'Antal användare')</th></tr></thead>
        <tbody>
"@

foreach ($s in $skuBreakdown) {
        $reportHtml += "    <tr><td>$($s.Sku)</td><td>$($s.Users)</td></tr>`n"
}

$reportHtml += @"
        </tbody>
    </table>
    </div>
</section>

<section class="card-section">
        <h2>$(Get-LocalizedText "Licensed users (showing $($licensedUsers.Count) users)" "Licensierade användare (visar $($licensedUsers.Count) användare)")</h2>
    <div class="table-wrapper">
        <table>
                <thead>
                        <tr>
                                <th>$(Get-LocalizedText 'Name' 'Namn')</th>
                                <th>$(Get-LocalizedText 'UPN' 'UPN')</th>
                                <th>$(Get-LocalizedText 'Email' 'E-post')</th>
                                <th>$(Get-LocalizedText 'Last login' 'Senaste inloggning')</th>
                                <th>$(Get-LocalizedText 'Action' 'Åtgärd')</th>
                                <th>$(Get-LocalizedText 'Licenses' 'Licenser')</th>
                                <th>$(Get-LocalizedText 'Count' 'Antal')</th>
                        </tr>
                </thead>
        <tbody>
"@

foreach ($u in $licensedUsers) {
        $name = [System.Web.HttpUtility]::HtmlEncode($u.DisplayName)
        $upn = [System.Web.HttpUtility]::HtmlEncode($u.UserPrincipalName)
        $mail = [System.Web.HttpUtility]::HtmlEncode($u.Mail)
        $lics = [System.Web.HttpUtility]::HtmlEncode($u.Licenses)
        $lastLogin = [System.Web.HttpUtility]::HtmlEncode($u.LastLoginDisplay)
        $actionLabel = [System.Web.HttpUtility]::HtmlEncode($u.LifecycleStatusLabel)
        $rowClassAttr = if ($u.LifecycleStatus -eq 'keep') { 'row-keep' } else { 'row-delete' }
        $badgeClassAttr = if ($u.LifecycleStatus -eq 'keep') { 'badge keep' } else { 'badge delete' }
        $keepLabelEsc = [System.Web.HttpUtility]::HtmlEncode($keepLabel)
        $deleteLabelEsc = [System.Web.HttpUtility]::HtmlEncode($deleteLabel)
        $statusValue = [System.Web.HttpUtility]::HtmlEncode($u.LifecycleStatus)
        $userIdentifier = if ($u.UserPrincipalName) { $u.UserPrincipalName } elseif ($u.Mail) { $u.Mail } else { $u.DisplayName }
        $userIdentifierEsc = [System.Web.HttpUtility]::HtmlEncode($userIdentifier)
        $cnt = $u.LicenseCount
        $reportHtml += "    <tr class='$rowClassAttr' data-status='$statusValue' data-keep-label='$keepLabelEsc' data-delete-label='$deleteLabelEsc' data-user-id='$userIdentifierEsc'><td>$name</td><td>$upn</td><td>$mail</td><td>$lastLogin</td><td><span class='$badgeClassAttr status-toggle'>$actionLabel</span></td><td>$lics</td><td>$cnt</td></tr>`n"
}

$reportHtml += @"
                </tbody>
        </table>
        </div>
</section>

<footer>
        <div class="small">$(Get-LocalizedText 'Report generated via Microsoft Graph PowerShell' 'Rapport genererad via Microsoft Graph PowerShell')</div>
</footer>
</div>
<script>
(function(){
        const btn = document.getElementById('legacyPrintButton');
        if(btn) {
                btn.addEventListener('click', function(){
                        document.body.classList.add('legacy-mode');
                        setTimeout(function(){
                                window.print();
                                document.body.classList.remove('legacy-mode');
                        }, 50);
                });
        }

        const keepMetric = document.getElementById('metric-keep');
        const deleteMetric = document.getElementById('metric-delete');

        function applyStatus(row, status) {
                if(!row) { return; }
                const normalized = status === 'keep' ? 'keep' : 'delete';
                row.dataset.status = normalized;
                row.classList.remove('row-keep','row-delete');
                row.classList.add(normalized === 'keep' ? 'row-keep' : 'row-delete');
                const badge = row.querySelector('.status-toggle');
                if (badge) {
                        badge.classList.remove('keep','delete');
                        badge.classList.add(normalized === 'keep' ? 'keep' : 'delete');
                        const keepLabel = row.dataset.keepLabel || 'Keep';
                        const deleteLabel = row.dataset.deleteLabel || 'Delete';
                        badge.textContent = normalized === 'keep' ? keepLabel : deleteLabel;
                }
        }

        function recalcSummary(){
                const rows = document.querySelectorAll('tbody tr[data-status]');
                let keepCount = 0;
                let deleteCount = 0;
                rows.forEach(function(row){
                        if (row.dataset.status === 'keep') { keepCount++; }
                        else { deleteCount++; }
                });
                if (keepMetric) { keepMetric.textContent = keepCount; }
                if (deleteMetric) { deleteMetric.textContent = deleteCount; }
        }

        document.querySelectorAll('tbody tr[data-status]').forEach(function(row){
                const badge = row.querySelector('.status-toggle');
                if (!badge) { return; }
                badge.addEventListener('click', function(){
                        const nextStatus = row.dataset.status === 'keep' ? 'delete' : 'keep';
                        applyStatus(row, nextStatus);
                        recalcSummary();
                });
        });
        recalcSummary();
})();
</script>
</body>
</html>
"@

# Save HTML to temp
$tempHtml = [System.IO.Path]::GetTempFileName() + ".html"
Set-Content -Path $tempHtml -Value $reportHtml -Encoding UTF8
if ($timestampFolder) {
        $archivedHtmlPath = Join-Path $timestampFolder "M365-LicensedUsers.html"
        Copy-Item -Path $tempHtml -Destination $archivedHtmlPath -Force
} else {
        $archivedHtmlPath = [System.IO.Path]::ChangeExtension($OutputPdfPath, ".html")
        Copy-Item -Path $tempHtml -Destination $archivedHtmlPath -Force
}

# PDF conversion helpers
function Install-WkHtmlToPdf {
        Write-Host "wkhtmltopdf not found. Attempting to install..." -ForegroundColor Yellow
        try {
                winget install wkhtmltopdf.wkhtmltox --silent --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                        Write-Host "wkhtmltopdf installed successfully. Please restart PowerShell and run the script again." -ForegroundColor Green
                        return $true
                } else {
                        Write-Host "Failed to install wkhtmltopdf automatically." -ForegroundColor Yellow
                        return $false
                }
        } catch {
                Write-Host "Error installing wkhtmltopdf: $_" -ForegroundColor Yellow
                return $false
        }
}

function Convert-With-WkHtml {
        param($htmlPath, $pdfPath)
        $wk = (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue).Source
        if (-not $wk) { 
                # Try to install it
                if (Install-WkHtmlToPdf) {
                        # Refresh PATH
                        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
                        $wk = (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue).Source
                        if (-not $wk) { return $false }
                } else {
                        return $false
                }
        }
        & $wk "--enable-local-file-access" "--print-media-type" "--margin-top" "15mm" "--margin-bottom" "15mm" "--margin-left" "15mm" "--margin-right" "15mm" "--page-size" "A4" $htmlPath $pdfPath
        return (Test-Path $pdfPath)
}

function Convert-With-Chrome {
        param($htmlPath, $pdfPath)
        # Try Edge then Chrome
        $candidates = @(
                "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
        ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

        if (-not $candidates) {
                return $false
        }

        $exe = $candidates

        # Use a unique temp file for output then copy
        $tempPdf = Join-Path $env:TEMP ("m365-report-" + [guid]::NewGuid().ToString() + ".pdf")
        if (Test-Path $tempPdf) { Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue }

        $uri = (New-Object System.Uri((Resolve-Path $htmlPath))).AbsoluteUri

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $exe
        $psi.Arguments = "--headless --disable-gpu --print-to-pdf=`"$tempPdf`" `"$uri`""
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardError = $true

        $p = [System.Diagnostics.Process]::Start($psi)
        if (-not $p.WaitForExit(20000)) {
                try { $p.Kill() } catch { }
                return $false
        }

        if (-not (Test-Path $tempPdf)) {
                return $false
        }

        Copy-Item $tempPdf $pdfPath -Force
        Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
        return (Test-Path $pdfPath)
}

# Try conversion
Write-Host "Converting HTML to PDF..." -ForegroundColor Cyan
$pdfCreated = $false

# Ensure output directory exists
$outDir = Split-Path -Path $OutputPdfPath -Parent
if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

if ($UseWkHtml) {
        if (Convert-With-WkHtml -htmlPath $tempHtml -pdfPath $OutputPdfPath) {
                $pdfCreated = $true
                Write-Host "PDF created using wkhtmltopdf: $OutputPdfPath" -ForegroundColor Green
        }
        else {
                Write-Host "wkhtmltopdf conversion failed; falling back to Edge/Chrome." -ForegroundColor Yellow
        }
}
else {
        Write-Host "Skipping wkhtmltopdf (enable with -UseWkHtml)." -ForegroundColor Yellow
}

if (-not $pdfCreated) {
        if (Convert-With-Chrome -htmlPath $tempHtml -pdfPath $OutputPdfPath) {
                $pdfCreated = $true
                Write-Host "PDF created using Headless Edge/Chrome: $OutputPdfPath" -ForegroundColor Green
        }
}

if ($pdfCreated -and -not (Test-Path $OutputPdfPath)) {
        Write-Host "Expected PDF at $OutputPdfPath but the file was not found after conversion." -ForegroundColor Yellow
        $pdfCreated = $false
}

if (-not $pdfCreated) {
        # fallback: save HTML next to desired PDF with .html extension
        $fallbackHtml = $archivedHtmlPath
        Write-Host "Could not convert to PDF automatically." -ForegroundColor Yellow
        Write-Host "HTML report saved to: $fallbackHtml" -ForegroundColor Green
        Write-Host "`nTo convert to PDF manually:" -ForegroundColor Cyan
        Write-Host "  1. Open the HTML file in Edge/Chrome" -ForegroundColor Cyan
        Write-Host "  2. Press Ctrl+P to print" -ForegroundColor Cyan
        Write-Host "  3. Select 'Save as PDF' as the printer" -ForegroundColor Cyan
        Write-Host "`nOr install wkhtmltopdf from: https://wkhtmltopdf.org/downloads.html" -ForegroundColor Cyan
}
elseif ($timestampFolder) {
        $outputDirRaw = Split-Path -Path $OutputPdfPath -Parent
        if (-not $outputDirRaw) { $outputDirRaw = (Get-Location).Path }
        try {
                $outputDirResolved = [System.IO.Path]::GetFullPath($outputDirRaw)
        } catch {
                $outputDirResolved = $outputDirRaw
        }
        try {
                $timestampResolved = [System.IO.Path]::GetFullPath($timestampFolder)
        } catch {
                $timestampResolved = $timestampFolder
        }
        if ($outputDirResolved -ne $timestampResolved) {
                $archivePdfPath = Join-Path $timestampFolder (Split-Path -Leaf $OutputPdfPath)
                Copy-Item -Path $OutputPdfPath -Destination $archivePdfPath -Force
                Write-Host "Archived PDF copy: $archivePdfPath" -ForegroundColor Cyan
        }
}

# cleanup temp html (keep fallback copy if used)
if (Test-Path $tempHtml) { Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue }

# Disconnect
Disconnect-MgGraph

if ($pdfCreated) { exit 0 } else { exit 1 }