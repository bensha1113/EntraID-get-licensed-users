<#
Copyright (c) 2025 Bensha1113
Licensed under the MIT License. See LICENSE or README for details.

.SYNOPSIS
        Generates a localized Microsoft 365 licensed-user inventory with interactive HTML, lifecycle KPIs, and optional PDF output.

.DESCRIPTION
        - Uses Microsoft Graph PowerShell (Microsoft.Graph) to enumerate users, subscribed SKUs, and directory subscription metadata.
        - Correlates recent sign-ins to classify each user into Keep / Review / Delete buckets with badge toggles and bulk actions.
        - Emits a single-page web dashboard (cards, donut charts, filters, bulk exports) plus legacy-print mode and optional PDF conversion
          via wkhtmltopdf or Edge/Chrome headless printing.
        - Supports 14 built-in languages through JSON translation packs under `translations/` or via a custom directory.
        - Archives every run under `Documents\User Reports\<tenant-domain>\<timestamp>` to keep compliance reviewers happy.
        - Requires delegated Microsoft Graph scopes: User.Read.All, Directory.Read.All, AuditLog.Read.All (sign-in lookups can be skipped).

.USAGE
        # Default HTML + PDF archived per-tenant
        pwsh .\Get-licensed-users.ps1

        # Swedish labels, custom inactivity horizon, and manual overrides
        pwsh .\Get-licensed-users.ps1 -Language Swedish -InactiveThresholdDays 60 -DecisionOverrideCsvPath .\overrides.csv

        # Force wkhtmltopdf and emit a PDF to a shared folder
        pwsh .\Get-licensed-users.ps1 -UseWkHtml -OutputPdfPath "C:\Reports\M365-LicensedUsers.pdf"

.PARAMETER OutputPdfPath
        Optional absolute PDF destination. When omitted the PDF (if created) is stored in the timestamped archive folder and
        a copy is made there even when a custom path is supplied.

.PARAMETER Language
        Display language for the HTML/PDF. Valid values: English, Mandarin, Hindi, Spanish, Arabic, Bengali, French, Portuguese,
        Russian, Urdu, Swedish, Norwegian, Danish, German. Defaults to English.

.PARAMETER UseWkHtml
        Switch that forces wkhtmltopdf usage even if Edge/Chrome headless printing is available.

.PARAMETER SkipSignInLookup
        Skip pulling `Get-MgAuditLogSignIn` data when your role lacks AuditLog.Read.All or you want faster dry-runs; lifecycle
        buckets will rely on overrides only.

.PARAMETER InactiveThresholdDays
        Number of days without a sign-in before a user is auto-classified as Delete. Default is 90; accepted range 1-3650.

.PARAMETER DecisionOverrideCsvPath
        Path to a CSV with UPN/Email + Keep/Review/Delete instructions to override the automated lifecycle decision.

.PARAMETER TranslationsDirectory
        Alternate directory that contains `<Language>.json` translation packs. When omitted the `translations` folder next to the script is used.

.NOTES
        - The script auto-installs Microsoft.Graph modules if missing and prompts for Graph consent on first run.
        - Refer to README.md for screenshots, documentation assets, and publishing guidance.
#>

param(
        [string]$OutputPdfPath,
        [ValidateSet("English", "Mandarin", "Hindi", "Spanish", "Arabic", "Bengali", "French", "Portuguese", "Russian", "Urdu", "Swedish", "Norwegian", "Danish", "German")]
        [string]$Language = "English",
        [switch]$UseWkHtml,
        [switch]$SkipSignInLookup,
        [ValidateRange(1, 3650)]
        [int]$InactiveThresholdDays = 90,
        [string]$DecisionOverrideCsvPath,
        [string]$TranslationsDirectory
)

# Resolve report language directly from parameter (defaults to English)
$script:ReportLanguage = $Language

if (-not $TranslationsDirectory -or [string]::IsNullOrWhiteSpace($TranslationsDirectory)) {
        if ($PSScriptRoot) {
                $TranslationsDirectory = Join-Path $PSScriptRoot 'translations'
        }
        else {
                $TranslationsDirectory = Join-Path (Get-Location) 'translations'
        }
}
$script:TranslationsDirectory = $TranslationsDirectory
$script:TranslationStore = @{}

function New-UnicodeGlyph {
        param([int[]]$CodePoints)
        if (-not $CodePoints) { return '' }
        $builder = New-Object System.Text.StringBuilder
        foreach ($cp in $CodePoints) {
                if ($cp -lt 0) { continue }
                $builder.Append([System.Char]::ConvertFromUtf32($cp)) | Out-Null
        }
        return $builder.ToString()
}

function Initialize-TranslationStore {
        param(
                [string]$Language,
                [string]$Directory
        )

        $script:TranslationStore = @{}
        if (-not $Language -or $Language -eq 'English') { return }
        if (-not $Directory -or -not (Test-Path -LiteralPath $Directory)) { return }
        try {
                $resolvedDirectory = (Resolve-Path -LiteralPath $Directory).Path
        }
        catch {
                Write-Warning "Could not resolve translations directory '$Directory': $_"
                return
        }

        $jsonPath = Join-Path $resolvedDirectory ("$Language.json")
        $psd1Path = Join-Path $resolvedDirectory ("$Language.psd1")

        if (Test-Path -LiteralPath $jsonPath) {
                try {
                        $jsonContent = Get-Content -LiteralPath $jsonPath -Raw -ErrorAction Stop
                        $script:TranslationStore = ConvertFrom-Json -InputObject $jsonContent -AsHashtable
                        return
                }
                catch {
                        Write-Warning "Failed to load ${jsonPath}: $_"
                }
        }

        if (Test-Path -LiteralPath $psd1Path) {
                try {
                        $script:TranslationStore = Import-PowerShellDataFile -Path $psd1Path
                        return
                }
                catch {
                        Write-Warning "Failed to load ${psd1Path}: $_"
                }
        }
}

function Resolve-LocalizedString {
        param(
                [Parameter(Mandatory)][string]$English,
                [string]$Swedish
        )

        $lang = $script:ReportLanguage
        if (-not $lang) { return $English }

        if ($script:TranslationStore -and $script:TranslationStore.ContainsKey($English)) {
                $resolved = $script:TranslationStore[$English]
                if ($null -ne $resolved -and $resolved -ne '') { return [string]$resolved }
        }

        if ($lang -eq 'Swedish' -and $Swedish) { return $Swedish }
        return $English
}

# Helper to render localized labels so the layout stays concise while supporting ENG/SWE
function Get-LocalizedText {
        param(
                [Parameter(Mandatory)] [string]$English,
                [Parameter(Mandatory)] [string]$Swedish
        )
        return (Resolve-LocalizedString -English $English -Swedish $Swedish)
}

function Get-LocalizedInlineText {
        param(
                [Parameter(Mandatory)] [string]$English,
                [Parameter(Mandatory)] [string]$Swedish
        )
        return (Resolve-LocalizedString -English $English -Swedish $Swedish)
}

Initialize-TranslationStore -Language $script:ReportLanguage -Directory $script:TranslationsDirectory

function Get-InferredBillingCadenceFromDates {
        param([Nullable[datetime]]$StartDate, [Nullable[datetime]]$EndDate)
        if (-not $StartDate -or -not $EndDate) { return $null }
        $days = [math]::Abs(($EndDate - $StartDate).TotalDays)
        if ($days -ge 300 -and $days -le 400) { return 'P1Y' }
        if ($days -ge 25 -and $days -le 40) { return 'P1M' }
        if ($days -gt 0) { return "P$([int][math]::Round($days))D" }
        return $null
}

function Install-ModuleIfMissing {
        param(
                [Parameter(Mandatory)][string]$Name,
                [switch]$AllowClobber
        )

        $alreadyAvailable = Get-Module -ListAvailable -Name $Name
        if ($alreadyAvailable) { return }

        Write-Host "Module '$Name' not found locally. Attempting install..." -ForegroundColor Yellow
        $installParams = @{
                Name        = $Name
                Scope       = 'CurrentUser'
                Force       = $true
                ErrorAction = 'Stop'
        }
        if ($AllowClobber) { $installParams.AllowClobber = $true }
        try {
                Install-Module @installParams
                Write-Host "Module '$Name' installed successfully." -ForegroundColor Green
        }
        catch {
                Write-Warning "Failed to install module '$Name': $_"
                throw
        }
}

function Get-LicenseNameMap {
        param([string]$CatalogUrl)

        $map = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
        if ([string]::IsNullOrWhiteSpace($CatalogUrl)) { return $map }

        $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ("license-catalog-" + [guid]::NewGuid().ToString() + ".csv")
        try {
                Invoke-WebRequest -Uri $CatalogUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop | Out-Null
                $rows = Import-Csv -Path $tempFile -ErrorAction Stop
                foreach ($row in $rows) {
                        if (-not $row) { continue }
                        $skuCandidate = $null
                        foreach ($prop in 'SkuPartNumber','String_Id','Product_Id','GUID') {
                                if ($row.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace($row.$prop)) {
                                        $skuCandidate = ($row.$prop).ToString().Trim()
                                        break
                                }
                        }
                        if (-not $skuCandidate) { continue }
                        $friendly = $null
                        foreach ($prop in 'Product_Display_Name','Service_Plan_Name') {
                                if ($row.PSObject.Properties[$prop] -and -not [string]::IsNullOrWhiteSpace($row.$prop)) {
                                        $friendly = ($row.$prop).ToString().Trim()
                                        break
                                }
                        }
                        if (-not $friendly) { continue }
                        if ($map.ContainsKey($skuCandidate)) { continue }
                        $map[$skuCandidate] = $friendly
                }
        }
        catch {
                Write-Warning "Could not load license catalog: $_"
        }
        finally {
                if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
        }
        return $map
}

function Get-NormalizedSkuId {
        param([Parameter(Mandatory)]$SkuId)
        if ($null -eq $SkuId) { return $null }
        if ($SkuId -is [guid]) { return $SkuId.ToString() }
        $text = "$SkuId"
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $trimmed = $text.Trim('{} ').Trim()
        $parsed = [guid]::Empty
        if ([guid]::TryParse($trimmed, [ref]$parsed)) { return $parsed.ToString() }
        return $trimmed
}

function Get-BillingCycleLabel {
        param([string]$RawCycle)
        if ([string]::IsNullOrWhiteSpace($RawCycle)) { return (Get-LocalizedText 'Unknown cadence' 'Okänt intervall') }
        $normalized = $RawCycle.Trim().ToUpperInvariant()
        switch -Regex ($normalized) {
                '^(P1Y|ANNUAL|YEARLY)$' { return (Get-LocalizedText 'Annual' 'Årsvis') }
                '^(P1M|MONTHLY)$' { return (Get-LocalizedText 'Monthly' 'Månadsvis') }
                '^(P1W|WEEKLY)$' { return (Get-LocalizedText 'Weekly' 'Veckovis') }
                '^(P1D|DAILY)$' { return (Get-LocalizedText 'Daily' 'Dagligen') }
                '^P(\d+)D$' {
                        $days = [int]$Matches[1]
                        return (Get-LocalizedText "Every $days days" "Var $days dag")
                }
                default { return $RawCycle }
        }
}

function Get-UserDocumentsFolder {
        $documents = [Environment]::GetFolderPath('MyDocuments')
        if ($documents -and (Test-Path $documents)) { return $documents }
        $fallback = Join-Path $HOME 'Documents'
        if (Test-Path $fallback) { return $fallback }
        return $HOME
}

function Invoke-GraphPagedRequest {
        param(
                [Parameter(Mandatory)][string]$Uri,
                [ValidateSet('GET','POST')][string]$Method = 'GET',
                $Body,
                [hashtable]$Headers
        )

        $results = @()
        $nextLink = $Uri
        while ($nextLink) {
                # Follow @odata.nextLink so large tenants do not silently drop later pages.
                $requestParams = @{ Method = $Method; Uri = $nextLink; OutputType = 'PSObject' }
                if ($Body) { $requestParams.Body = $Body }
                if ($Headers) { $requestParams.Headers = $Headers }
                $response = Invoke-MgGraphRequest @requestParams
                if ($response.value) { $results += $response.value }
                elseif ($response) { $results += $response }
                if ($response.'@odata.nextLink') { $nextLink = $response.'@odata.nextLink' }
                else { $nextLink = $null }
        }
        return $results
}

$script:AdminRoleIconRules = @(
        @{ Pattern = 'Global|Company Administrator'; Icon = New-UnicodeGlyph 0x1F310 }
        @{ Pattern = 'Billing'; Icon = New-UnicodeGlyph 0x1F4B3 }
        @{ Pattern = 'SharePoint'; Icon = New-UnicodeGlyph 0x1F4C1 }
        @{ Pattern = 'Exchange'; Icon = (New-UnicodeGlyph 0x2709 0xFE0F) }
        @{ Pattern = 'Teams|Skype'; Icon = New-UnicodeGlyph 0x1F4AC }
        @{ Pattern = 'Security'; Icon = (New-UnicodeGlyph 0x1F6E1 0xFE0F) }
        @{ Pattern = 'Compliance|eDiscovery'; Icon = New-UnicodeGlyph 0x1F4CB }
        @{ Pattern = 'Helpdesk|Support|Service'; Icon = New-UnicodeGlyph 0x1F198 }
        @{ Pattern = 'Application|App'; Icon = (New-UnicodeGlyph 0x2699 0xFE0F) }
        @{ Pattern = 'Password'; Icon = New-UnicodeGlyph 0x1F511 }
        @{ Pattern = 'Device|Endpoint|Intune'; Icon = New-UnicodeGlyph 0x1F4BB }
)

function Get-AdminRoleIconSymbol {
        param([string]$RoleName)
        if ([string]::IsNullOrWhiteSpace($RoleName)) { return (New-UnicodeGlyph 0x1F6E1 0xFE0F) }
        foreach ($rule in $script:AdminRoleIconRules) {
                if ($RoleName -match $rule.Pattern) { return $rule.Icon }
        }
        return (New-UnicodeGlyph 0x1F6E1 0xFE0F)
}

# Ensure Microsoft Graph module
Write-Host "Checking for Microsoft.Graph module..." -ForegroundColor Cyan
Install-ModuleIfMissing -Name Microsoft.Graph

# Import module
Write-Host "Importing Microsoft.Graph modules..." -ForegroundColor Cyan
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
Import-Module Microsoft.Graph.Reports -ErrorAction SilentlyContinue
Write-Host "Modules imported successfully." -ForegroundColor Green

# Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
$scopes = @("User.Read.All", "Directory.Read.All", "AuditLog.Read.All")
Connect-MgGraph -Scopes $scopes

# Retrieve SKU and subscription inventory
Write-Host "Retrieving subscribed SKUs..." -ForegroundColor Cyan
$skuList = @()
try {
        $skuList = Get-MgSubscribedSku -All
}
catch {
        Write-Warning "Could not retrieve subscribed SKU inventory: $_"
}
$skuMap = @{}
foreach ($sku in $skuList) {
        if ($sku.SkuId) {
                $skuMap[$sku.SkuId] = $sku.SkuPartNumber
        }
}
Write-Host "Retrieved $($skuList.Count) subscribed SKU records" -ForegroundColor Cyan

$licenseCatalogUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$licenseNameMap = Get-LicenseNameMap -CatalogUrl $licenseCatalogUrl
if ($licenseNameMap.Count -gt 0) {
        Write-Host "Loaded $($licenseNameMap.Count) friendly license names" -ForegroundColor Cyan
}
else {
        Write-Host "License catalog could not be loaded; falling back to SKU identifiers" -ForegroundColor Yellow
}

$skuInventoryRows = @()
foreach ($sku in $skuList) {
        if (-not $sku) { continue }
        $normalizedSkuId = Get-NormalizedSkuId -SkuId $sku.SkuId
        $skuPartNumber = $sku.SkuPartNumber
        $friendlyName = $null
        if ($skuPartNumber -and $licenseNameMap.ContainsKey($skuPartNumber)) {
                $friendlyName = $licenseNameMap[$skuPartNumber]
        }
        elseif ($sku.SkuPartDescription) {
                $friendlyName = $sku.SkuPartDescription
        }
        elseif ($skuPartNumber) {
                $friendlyName = $skuPartNumber
        }
        else {
                $friendlyName = if ($normalizedSkuId) { "SKU $normalizedSkuId" } else { 'Unknown SKU' }
        }

        $prepaidUnits = $null
        $enabledUnitsRaw = $null
        $includedUnitsRaw = $null
        if ($null -ne $sku.PrepaidUnits) {
                $enabledUnitsRaw = $sku.PrepaidUnits.Enabled
                $includedUnitsRaw = $sku.PrepaidUnits.Included
        }
        if ($null -ne $enabledUnitsRaw) {
                $prepaidUnits = [int]$enabledUnitsRaw
        }
        elseif ($null -ne $includedUnitsRaw) {
                $prepaidUnits = [int]$includedUnitsRaw
        }
        $consumedUnitsRaw = $sku.ConsumedUnits
        $consumedUnits = if ($null -ne $consumedUnitsRaw) { [int]$consumedUnitsRaw } else { $null }
        $usagePercent = $null
        if ($prepaidUnits -and $prepaidUnits -gt 0 -and $null -ne $consumedUnits) {
                $usagePercent = [math]::Round(($consumedUnits / $prepaidUnits) * 100, 1)
        }

        $skuInventoryRows += [PSCustomObject]@{
                SkuId = $normalizedSkuId
                SkuPartNumber = $skuPartNumber
                FriendlyName = $friendlyName
                CapabilityStatus = $sku.CapabilityStatus
                AppliesTo = $sku.AppliesTo
                PurchasedUnits = $prepaidUnits
                ConsumedUnits = $consumedUnits
                UsagePercent = $usagePercent
                BillingCycle = $null
                BillingCycleRaw = $null
                TermStart = $null
                TermEnd = $null
                DaysRemaining = $null
                AutoRenew = $null
                SubscriptionChannel = $null
                SubscriptionStatus = $sku.CapabilityStatus
                IsTrial = $false
        }
}

$subscriptionMetadataFallbackMessage = Get-LocalizedText 'Detailed subscription metadata is unavailable right now (Directory.Subscriptions permission may be required). Showing aggregated SKU totals instead.' 'Detaljerad abonnemangsmetadata är inte tillgänglig just nu (kräver behörigheten Directory.Subscriptions). Visar i stället aggregerade SKU-summor.'
$directorySubscriptionRecords = @()
$directorySubscriptionAvailable = $false
try {
        Write-Host "Retrieving directory subscription metadata..." -ForegroundColor Cyan
        $subscriptionEndpoint = "https://graph.microsoft.com/beta/directory/subscriptions"
        $directorySubscriptionRecords = Invoke-GraphPagedRequest -Uri $subscriptionEndpoint
        if ($directorySubscriptionRecords.Count -gt 0) {
                $directorySubscriptionAvailable = $true
                Write-Host "Captured $($directorySubscriptionRecords.Count) directory subscription entries" -ForegroundColor Cyan
        }
        else {
                Write-Host "Directory subscription metadata returned no rows; using SKU summary only." -ForegroundColor Yellow
        }
}
catch {
        Write-Warning "Could not retrieve directory subscription metadata: $_"
        $directorySubscriptionRecords = @()
}

$skuInventoryLookup = @{}
foreach ($row in $skuInventoryRows) {
        if ($row.SkuId) {
                $skuInventoryLookup[$row.SkuId] = $row
        }
}

$subscriptionDisplayRows = @()
foreach ($subscription in $directorySubscriptionRecords) {
        $rawSubscriptionSku = $subscription.skuId
        if (-not $rawSubscriptionSku -and $subscription.productSkuId) { $rawSubscriptionSku = $subscription.productSkuId }
        $normalizedSubscriptionSku = Get-NormalizedSkuId -SkuId $rawSubscriptionSku
        $associatedSku = $null
        if ($normalizedSubscriptionSku -and $skuInventoryLookup.ContainsKey($normalizedSubscriptionSku)) {
                $associatedSku = $skuInventoryLookup[$normalizedSubscriptionSku]
        }
        $skuPartNumber = $null
        if ($normalizedSubscriptionSku -and $skuMap.ContainsKey($normalizedSubscriptionSku)) {
                $skuPartNumber = $skuMap[$normalizedSubscriptionSku]
        }
        if (-not $skuPartNumber -and $associatedSku) { $skuPartNumber = $associatedSku.SkuPartNumber }
        $friendlyName = $null
        if ($skuPartNumber -and $licenseNameMap.ContainsKey($skuPartNumber)) {
                $friendlyName = $licenseNameMap[$skuPartNumber]
        }
        if (-not $friendlyName -and $associatedSku) { $friendlyName = $associatedSku.FriendlyName }
        if (-not $friendlyName -and $subscription.displayName) { $friendlyName = $subscription.displayName }
        if (-not $friendlyName -and $skuPartNumber) { $friendlyName = $skuPartNumber }
        if (-not $friendlyName) { $friendlyName = if ($normalizedSubscriptionSku) { "SKU $normalizedSubscriptionSku" } else { 'Unmapped subscription' } }

        $billingCycleRaw = $null
        foreach ($prop in @('billingCycle','billingTerm','subscriptionTerm')) {
                $candidate = $subscription.$prop
                if ($null -ne $candidate -and $candidate -ne '') { $billingCycleRaw = $candidate; break }
        }
        if (-not $billingCycleRaw -and $subscription.termInformation) {
                if ($subscription.termInformation.billingCycle) { $billingCycleRaw = $subscription.termInformation.billingCycle }
                elseif ($subscription.termInformation.termDuration) { $billingCycleRaw = $subscription.termInformation.termDuration }
        }
        $termStartDate = $null
        foreach ($prop in @('termStartDateTime','startDateTime','startDate','createdDateTime')) {
                $candidate = $subscription.$prop
                if ($candidate) { try { $termStartDate = [datetime]$candidate } catch { } }
                if ($termStartDate) { break }
        }
        $termEndDate = $null
        foreach ($prop in @('termEndDateTime','subscriptionEndDateTime','nextLifecycleDateTime','endDateTime','expirationDateTime')) {
                $candidate = $subscription.$prop
                if ($candidate) { try { $termEndDate = [datetime]$candidate } catch { } }
                if ($termEndDate) { break }
        }
                if (-not $billingCycleRaw) {
                        # Older subscriptions often omit billingCycle; infer it from the term window so charts stay meaningful.
            $inferredCycle = Get-InferredBillingCadenceFromDates -StartDate $termStartDate -EndDate $termEndDate
            if ($inferredCycle) { $billingCycleRaw = $inferredCycle }
        }
        $billingLabel = Get-BillingCycleLabel -RawCycle $billingCycleRaw

        $purchasedUnitsRaw = $null
        foreach ($prop in @('totalLicenses','totalLicenseCount','licenseCount','quantity','totalUnits')) {
                $candidate = $subscription.$prop
                if ($null -ne $candidate) { $purchasedUnitsRaw = $candidate; break }
        }
        if ($null -eq $purchasedUnitsRaw -and $associatedSku) { $purchasedUnitsRaw = $associatedSku.PurchasedUnits }
        $purchasedUnits = if ($null -ne $purchasedUnitsRaw -and $purchasedUnitsRaw -ne '') { [int]$purchasedUnitsRaw } else { $null }

        $consumedUnitsRaw = $null
        foreach ($prop in @('consumedUnits','usedLicenses','assignedLicenses','consumedLicenses','activeSeats')) {
                $candidate = $subscription.$prop
                if ($null -ne $candidate) { $consumedUnitsRaw = $candidate; break }
        }
        if ($null -eq $consumedUnitsRaw -and $associatedSku) { $consumedUnitsRaw = $associatedSku.ConsumedUnits }
        $consumedUnits = if ($null -ne $consumedUnitsRaw -and $consumedUnitsRaw -ne '') { [int]$consumedUnitsRaw } else { $null }

        $usagePercent = $null
        if ($purchasedUnits -and $purchasedUnits -gt 0 -and $null -ne $consumedUnits) {
                $usagePercent = [math]::Round(($consumedUnits / $purchasedUnits) * 100, 1)
        }
        elseif ($associatedSku) {
                $usagePercent = $associatedSku.UsagePercent
        }

        $autoRenew = $null
        foreach ($prop in @('autoRenewEnabled','autoRenew')) {
                if ($null -ne $subscription.$prop) { $autoRenew = [bool]$subscription.$prop; break }
        }
        if (-not $autoRenew -and $subscription.renewalOption) {
                $autoRenew = $subscription.renewalOption -match 'auto'
        }

        $channelLabel = $null
        foreach ($prop in @('subscriptionType','billingType','channel','offerType')) {
                $candidate = $subscription.$prop
                if ($candidate) { $channelLabel = $candidate; break }
        }
        if (-not $channelLabel -and $associatedSku) { $channelLabel = $associatedSku.SubscriptionChannel }
        if (-not $channelLabel) { $channelLabel = 'Unknown' }

        $statusLabel = $null
        foreach ($prop in @('status','lifecycleStatus','subscriptionStatus')) {
                $candidate = $subscription.$prop
                if ($candidate) { $statusLabel = $candidate; break }
        }
        if (-not $statusLabel -and $associatedSku) { $statusLabel = $associatedSku.SubscriptionStatus }

        $isTrial = $false
        if ($subscription.isTrial) { $isTrial = [bool]$subscription.isTrial }
        elseif ($subscription.offerType -and $subscription.offerType -match 'trial') { $isTrial = $true }
        elseif ($statusLabel -and $statusLabel -match 'trial') { $isTrial = $true }

        $daysRemaining = $null
        if ($termEndDate) { $daysRemaining = [math]::Round(($termEndDate - (Get-Date)).TotalDays, 0) }

        $subscriptionDisplayRows += [PSCustomObject]@{
                SkuId = $normalizedSubscriptionSku
                SkuPartNumber = $skuPartNumber
                FriendlyName = $friendlyName
                CapabilityStatus = $statusLabel
                AppliesTo = $subscription.AppliesTo
                PurchasedUnits = $purchasedUnits
                ConsumedUnits = $consumedUnits
                UsagePercent = if ($null -ne $usagePercent) { $usagePercent } else { $associatedSku.UsagePercent }
                BillingCycle = $billingLabel
                BillingCycleRaw = $billingCycleRaw
                TermStart = $termStartDate
                TermEnd = $termEndDate
                DaysRemaining = $daysRemaining
                AutoRenew = $autoRenew
                SubscriptionChannel = $channelLabel
                SubscriptionStatus = $statusLabel
                IsTrial = $isTrial
        }
}

if ($subscriptionDisplayRows.Count -gt 0) {
        $skuPurchaseTable = $subscriptionDisplayRows | Sort-Object -Property FriendlyName, BillingCycle
}
else {
        $skuPurchaseTable = $skuInventoryRows | Sort-Object -Property FriendlyName
}

$nextExpirySource = if ($subscriptionDisplayRows.Count -gt 0) { $subscriptionDisplayRows } else { $skuInventoryRows }
$nextExpiryDate = $null
if ($nextExpirySource) {
        $nextExpiryDate = ($nextExpirySource | Where-Object { $_.TermEnd } | Sort-Object -Property TermEnd | Select-Object -First 1).TermEnd
}
$nextExpiryDisplay = if ($nextExpiryDate) { $nextExpiryDate.ToString('yyyy-MM-dd') } else { Get-LocalizedText 'Unknown' 'Okänt' }
$aggregatedSkuNote = Get-LocalizedText 'Detailed subscription metadata is unavailable; showing aggregated SKU totals (monthly and yearly seats may appear combined).' 'Detaljerad prenumerationsmetadata är inte tillgänglig; visar aggregerade SKU-summor (månads- och årsplats kan kombineras).'
$subscriptionTableNote = if ($subscriptionDisplayRows.Count -eq 0) { $aggregatedSkuNote } else { $subscriptionNoteText }
$unknownLabel = Get-LocalizedText 'Unknown' 'Okänt'
$naLabel = Get-LocalizedText 'N/A' 'Saknas'
$trialLabel = Get-LocalizedText 'Trial' 'Testperiod'

# Get all users with assignedLicenses property
Write-Host "Retrieving all users (this may take a while)..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property "Id,AssignedLicenses,DisplayName,UserPrincipalName,Mail" -ConsistencyLevel eventual
Write-Host "Retrieved $($users.Count) users" -ForegroundColor Cyan

$userAdminRoleLookup = @{}
try {
        Write-Host "Retrieving directory role assignments..." -ForegroundColor Cyan
        $directoryRoles = Get-MgDirectoryRole -All -Property "Id,DisplayName"
        foreach ($role in $directoryRoles) {
                if (-not $role.Id) { continue }
                $roleName = if ($role.DisplayName) { $role.DisplayName } else { 'Directory role' }
                try {
                        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
                        foreach ($member in $members) {
                                $memberId = $member.Id
                                if (-not $memberId) { continue }
                                $odataType = $null
                                if ($member.PSObject.Properties["AdditionalProperties"]) {
                                        $odataType = $member.AdditionalProperties['@odata.type']
                                }
                                elseif ($member.PSObject.Properties["AdditionalData"]) {
                                        $odataType = $member.AdditionalData['@odata.type']
                                }
                                # Directory roles can include service principals; skip everything that is not a user object.
                                if ($odataType -and $odataType -ne '#microsoft.graph.user') { continue }
                                if (-not $odataType) {
                                        $typeNames = $member.PSObject.TypeNames
                                        if ($typeNames -and -not ($typeNames -contains 'Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser')) {
                                                continue
                                        }
                                }
                                if (-not $userAdminRoleLookup.ContainsKey($memberId)) {
                                        $userAdminRoleLookup[$memberId] = @()
                                }
                                if (-not ($userAdminRoleLookup[$memberId] | Where-Object { $_.Name -eq $roleName })) {
                                        $userAdminRoleLookup[$memberId] += [PSCustomObject]@{
                                                Name = $roleName
                                                Icon = Get-AdminRoleIconSymbol -RoleName $roleName
                                        }
                                }
                        }
                }
                catch {
                        Write-Warning "Could not read members for role ${roleName}: $_"
                }
        }
        if ($userAdminRoleLookup.Count -gt 0) {
                Write-Host "Mapped admin roles for $($userAdminRoleLookup.Count) users" -ForegroundColor Cyan
        }
        else {
                Write-Host "No admin role assignments detected." -ForegroundColor Yellow
        }
}
catch {
        Write-Warning "Could not retrieve directory roles: $_"
        $userAdminRoleLookup = @{}
}

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
                                # Normalize friendly terms into the three lifecycle buckets so CSV authors get flexibility without breaking UI logic.
                                switch -Regex ($rawDecision.ToString().Trim().ToLowerInvariant()) {
                                        '^(keep|retain|green|stay)$' { $decisionOverrides[$id] = 'keep'; continue }
                                        '^(delete|remove|drop|red)$' { $decisionOverrides[$id] = 'delete'; continue }
                                        '^(review|yellow|pending|hold)$' { $decisionOverrides[$id] = 'review'; continue }
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
$signInEvaluationEnabled = -not $SkipSignInLookup
if ($signInEvaluationEnabled) {
        # Stretch the audit query just enough to cover the inactivity threshold without asking Graph for year-long payloads.
        $signInLookbackDays = [math]::Min([math]::Max($InactiveThresholdDays, 30), 365)
        $logSince = (Get-Date).AddDays(-1 * $signInLookbackDays)
        $filterTimestamp = $logSince.ToUniversalTime().ToString("o")
        Write-Host "Retrieving sign-in logs since $filterTimestamp..." -ForegroundColor Cyan
        try {
                # Large single-window queries are prone to HttpClient timeouts, so slice the
                # lookback period into smaller chunks that the Graph service can fulfill quickly.
                $chunkDays = 30
                $chunkStart = $logSince
                $chunkEndCap = Get-Date
                $totalEvents = 0
                $totalChunks = 0
                while ($chunkStart -lt $chunkEndCap) {
                        $chunkEnd = $chunkStart.AddDays($chunkDays)
                        if ($chunkEnd -gt $chunkEndCap) { $chunkEnd = $chunkEndCap }
                        $chunkFilter = "createdDateTime ge $($chunkStart.ToUniversalTime().ToString('o')) and createdDateTime lt $($chunkEnd.ToUniversalTime().ToString('o'))"
                        $chunkLabel = "$(($chunkStart.ToUniversalTime()).ToString('yyyy-MM-dd')) -> $(($chunkEnd.ToUniversalTime()).ToString('yyyy-MM-dd'))"
                        Write-Host "  Fetching sign-ins for $chunkLabel" -ForegroundColor DarkGray
                        $chunkAttempt = 0
                        while ($true) {
                                try {
                                        $chunkLogs = Get-MgAuditLogSignIn -All -PageSize 200 -Filter $chunkFilter -Property userPrincipalName,createdDateTime -ErrorAction Stop
                                        break
                                }
                                catch {
                                        $chunkAttempt++
                                        if ($chunkAttempt -ge 3) { throw }
                                        $retryDelay = [math]::Min(5 * $chunkAttempt, 15)
                                        Write-Warning "Sign-in chunk $chunkLabel failed (attempt $chunkAttempt). Retrying in $retryDelay seconds..."
                                        Start-Sleep -Seconds $retryDelay
                                }
                        }
                        $chunkCount = 0
                        foreach ($entry in $chunkLogs) {
                                $upn = $entry.UserPrincipalName
                                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                                if (-not $entry.CreatedDateTime) { continue }
                                $entryTime = [datetime]$entry.CreatedDateTime
                                if (-not $signInLookup.ContainsKey($upn) -or $signInLookup[$upn] -lt $entryTime) {
                                        $signInLookup[$upn] = $entryTime
                                }
                                $chunkCount++
                        }
                        $totalEvents += $chunkCount
                        $totalChunks++
                        $chunkStart = $chunkEnd
                }
                Write-Host "Captured sign-in timestamps for $($signInLookup.Count) users ($totalEvents events across $totalChunks requests)" -ForegroundColor Cyan
        }
        catch {
                Write-Warning "Could not retrieve sign-in logs: $_"
        }
}
else {
        Write-Host "SkipSignInLookup set: skipping audit log retrieval and lifecycle scoring based on last login." -ForegroundColor Yellow
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
        } elseif (-not $signInEvaluationEnabled) {
                'keep'
        } elseif ($lastLogin -and $lastLogin -ge $inactiveCutoff) {
                'keep'
        } else {
                'review'
        }
        
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

        $adminRoles = @()
        if ($u.Id -and $userAdminRoleLookup.ContainsKey($u.Id)) {
                $adminRoles = $userAdminRoleLookup[$u.Id]
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
                LifecycleStatusLabel = switch ($status) {
                        'keep' { $keepLabel }
                        'delete' { $deleteLabel }
                        'review' { Get-LocalizedText 'Review' 'Granska' }
                        default { $status }
                }
                AdminRoles = $adminRoles
                HasAdminRights = ($adminRoles.Count -gt 0)
        }
} | Where-Object { $_.LicenseCount -gt 0 } | Sort-Object -Property DisplayName

Write-Host "Found $($licensedUsers.Count) licensed users" -ForegroundColor Cyan

# SKU breakdown (used by charts and tables)
$skuBreakdown = $licensedUsers |
        ForEach-Object { $_.Licenses -split ",\s*" } |
        Where-Object { $_ -ne "" } |
        Group-Object |
        Sort-Object Count -Descending |
        Select-Object @{n='Sku';e={$_.Name}}, @{n='Users';e={$_.Count}}

$totalUsers = $users.Count
$totalLicensed = $licensedUsers.Count
$totalUnlicensed = $totalUsers - $totalLicensed
$keepCount = ($licensedUsers | Where-Object { $_.LifecycleStatus -eq 'keep' }).Count
$reviewCount = ($licensedUsers | Where-Object { $_.LifecycleStatus -eq 'review' }).Count
$deleteCount = $totalLicensed - $keepCount - $reviewCount
$adminCount = ($licensedUsers | Where-Object { $_.HasAdminRights }).Count
$neverSignedInCount = ($licensedUsers | Where-Object { -not $_.LastLogin }).Count
if ($deleteCount -lt 0) { $deleteCount = 0 }
Write-Host "Admin accounts detected: $adminCount; Never signed-in (licensed) users: $neverSignedInCount" -ForegroundColor Cyan

$licensedTotalSafe = if ($totalLicensed -gt 0) { [double]$totalLicensed } else { 1.0 }
$lifecyclePercentKeepRounded = [math]::Round(($keepCount / $licensedTotalSafe) * 100, 0)
$adminPercentRounded = [math]::Round(($adminCount / $licensedTotalSafe) * 100, 0)
$nonAdminPercentRounded = [math]::Round((($totalLicensed - $adminCount) / $licensedTotalSafe) * 100, 0)
$topSkuLeader = if ($skuBreakdown.Count -gt 0) { $skuBreakdown | Select-Object -First 1 } else { $null }
$topSkuLeaderName = if ($topSkuLeader) { $topSkuLeader.Sku } else { (Get-LocalizedText 'No license data yet' 'Inga licensdata ännu') }
$topSkuLeaderCount = if ($topSkuLeader) { [int]$topSkuLeader.Users } else { 0 }
$topSkuLeaderPercentRounded = if ($topSkuLeader -and $totalLicensed -gt 0) { [math]::Round(($topSkuLeader.Users / $licensedTotalSafe) * 100, 0) } else { 0 }

$usagePulseLabel = Get-LocalizedText 'Usage pulse' 'Användningspuls'
$privExposureLabel = Get-LocalizedText 'Privilege exposure' 'Behörighetsexponering'
$licenseMomentumLabel = Get-LocalizedText 'License momentum' 'Licensmomentum'
$chartLiveNote = Get-LocalizedText 'Counts update instantly when you filter or adjust lifecycle badges.' 'Antal uppdateras direkt när du filtrerar eller ändrar livscykelbadges.'
$adminChartNote = Get-LocalizedText 'Track how many privileged accounts remain in view.' 'Följ hur många privilegierade konton som visas.'
$skuChartNote = Get-LocalizedText 'See where license assignments concentrate the most seats.' 'Se var licenstilldelningar koncentrerar flest platser.'
$shareLabel = Get-LocalizedText 'Share' 'Andel'
$seatsLabel = Get-LocalizedText 'Seats' 'Platser'
$adminsShortLabel = Get-LocalizedText 'Admins' 'Administratörer'
$nonAdminsShortLabel = Get-LocalizedText 'Non-admins' 'Icke-administratörer'
$topLicenseLabel = Get-LocalizedText 'Top license' 'Topplicens'
$noSkuDataLabel = Get-LocalizedText 'Not enough license data yet' 'Inte tillräckligt med licensdata ännu'

$topSkuItems = $skuBreakdown | Select-Object -First 5
$topSkuLabels = @()
$topSkuCounts = @()
foreach ($entry in $topSkuItems) {
        $topSkuLabels += $entry.Sku
        $topSkuCounts += $entry.Users
}
$subscriptionSuccessNote = Get-LocalizedText 'Renewal dates and lifecycle status come from directory/subscriptions. Data may be limited for certain offers.' 'Förnyelsedatum och livscykelstatus hämtas från directory/subscriptions. Data kan vara begränsad för vissa erbjudanden.'
$subscriptionNoteText = if ($directorySubscriptionAvailable -and $directorySubscriptionRecords.Count -gt 0) { $subscriptionSuccessNote } else { $subscriptionMetadataFallbackMessage }

# Build HTML
$now = $reportGeneratedAt.ToString("u")
$searchPlaceholder = Get-LocalizedInlineText 'Search name / email / license' 'Sök namn / e-post / licens'
$searchPlaceholderEsc = [System.Web.HttpUtility]::HtmlEncode($searchPlaceholder)
$groupByLabel = Get-LocalizedText 'Group by subscription type' 'Gruppera efter prenumerationstyp'
$groupByLabelEsc = [System.Web.HttpUtility]::HtmlEncode($groupByLabel)
$groupResetLabel = Get-LocalizedText 'Show original order' 'Visa ursprunglig ordning'
$groupResetLabelEsc = [System.Web.HttpUtility]::HtmlEncode($groupResetLabel)
$statusFilterAllLabel = Get-LocalizedInlineText 'All statuses' 'Alla statusar'
$statusFilterAllLabelEsc = [System.Web.HttpUtility]::HtmlEncode($statusFilterAllLabel)
$statusFilterToolbarLabel = Get-LocalizedInlineText 'Status filters' 'Statusfilter'
$statusFilterToolbarLabelEsc = [System.Web.HttpUtility]::HtmlEncode($statusFilterToolbarLabel)
$showAdminsOnlyLabel = Get-LocalizedInlineText 'Show admins only' 'Visa endast administratörer'
$showAdminsOnlyLabelEsc = [System.Web.HttpUtility]::HtmlEncode($showAdminsOnlyLabel)
$exportCsvLabel = Get-LocalizedInlineText 'Download filtered CSV' 'Ladda ner filtrerad CSV'
$exportCsvLabelEsc = [System.Web.HttpUtility]::HtmlEncode($exportCsvLabel)
$keepFilterLabel = Get-LocalizedInlineText 'Keep' 'Behåll'
$keepFilterLabelEsc = [System.Web.HttpUtility]::HtmlEncode($keepFilterLabel)
$reviewFilterLabel = Get-LocalizedInlineText 'Review' 'Granska'
$reviewFilterLabelEsc = [System.Web.HttpUtility]::HtmlEncode($reviewFilterLabel)
$deleteFilterLabel = Get-LocalizedInlineText 'Delete' 'Ta bort'
$deleteFilterLabelEsc = [System.Web.HttpUtility]::HtmlEncode($deleteFilterLabel)
$exportEmptyMessage = Get-LocalizedInlineText 'No rows match the current filters.' 'Inga rader matchar de aktuella filtren.'
$exportEmptyMessageEsc = [System.Web.HttpUtility]::HtmlEncode($exportEmptyMessage)
$visualInsightsTitle = Get-LocalizedText 'Visual insights' 'Visuella insikter'
$lifecycleChartTitle = Get-LocalizedText 'Lifecycle recommendations' 'Rekommendationer per status'
$adminChartTitle = Get-LocalizedText 'Admin exposure' 'Administratörsspridning'
$skuChartTitle = Get-LocalizedText 'Top license usage' 'Mest använda licenser'
$visualInsightsTitleEsc = [System.Web.HttpUtility]::HtmlEncode($visualInsightsTitle)
$lifecycleChartTitleEsc = [System.Web.HttpUtility]::HtmlEncode($lifecycleChartTitle)
$adminChartTitleEsc = [System.Web.HttpUtility]::HtmlEncode($adminChartTitle)
$skuChartTitleEsc = [System.Web.HttpUtility]::HtmlEncode($skuChartTitle)
$usagePulseLabelEsc = [System.Web.HttpUtility]::HtmlEncode($usagePulseLabel)
$privExposureLabelEsc = [System.Web.HttpUtility]::HtmlEncode($privExposureLabel)
$licenseMomentumLabelEsc = [System.Web.HttpUtility]::HtmlEncode($licenseMomentumLabel)
$chartLiveNoteEsc = [System.Web.HttpUtility]::HtmlEncode($chartLiveNote)
$adminChartNoteEsc = [System.Web.HttpUtility]::HtmlEncode($adminChartNote)
$skuChartNoteEsc = [System.Web.HttpUtility]::HtmlEncode($skuChartNote)
$shareLabelEsc = [System.Web.HttpUtility]::HtmlEncode($shareLabel)
$seatsLabelEsc = [System.Web.HttpUtility]::HtmlEncode($seatsLabel)
$adminsShortLabelEsc = [System.Web.HttpUtility]::HtmlEncode($adminsShortLabel)
$nonAdminsShortLabelEsc = [System.Web.HttpUtility]::HtmlEncode($nonAdminsShortLabel)
$tenantDomainEsc = [System.Web.HttpUtility]::HtmlEncode($tenantDomain)
$nowEsc = [System.Web.HttpUtility]::HtmlEncode($now)
$nextExpiryEsc = [System.Web.HttpUtility]::HtmlEncode($nextExpiryDisplay)
$visualInsightsSubtitle = Get-LocalizedText 'Dashboards react instantly to the roster filters below.' 'Instrumentpanelerna reagerar direkt på filtren nedan.'
$visualInsightsSubtitleEsc = [System.Web.HttpUtility]::HtmlEncode($visualInsightsSubtitle)
$noSkuDataLabelAttr = [System.Web.HttpUtility]::HtmlAttributeEncode($noSkuDataLabel)
$topSkuCenterValueText = if ($topSkuLeaderCount -gt 0 -and $totalLicensed -gt 0) { "$topSkuLeaderPercentRounded%" } elseif ($topSkuLeaderCount -gt 0) { [string]$topSkuLeaderCount } else { '—' }
$topSkuCenterValueEsc = [System.Web.HttpUtility]::HtmlEncode($topSkuCenterValueText)
$topSkuCenterLabelText = if ($topSkuLeaderName -and $topSkuLeaderCount -gt 0) { $topSkuLeaderName } else { $topLicenseLabel }
$topSkuCenterLabelEsc = [System.Web.HttpUtility]::HtmlEncode($topSkuCenterLabelText)
$topSkuLeaderTitleAttr = if ($topSkuLeaderName -and $topSkuLeaderCount -gt 0) { [System.Web.HttpUtility]::HtmlAttributeEncode($topSkuLeaderName) } else { '' }
$skuCenterLabelTitleAttr = if ($topSkuLeaderTitleAttr) { " title=`"$topSkuLeaderTitleAttr`"" } else { '' }
$csvFileName = "$safeTenantDomain-licensed-users.csv"
$csvFileNameEsc = [System.Web.HttpUtility]::HtmlEncode($csvFileName)
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
                padding: 20px 22px 40px;
                border-radius: 16px;
                box-shadow: 0 18px 38px -16px rgba(15, 23, 42, 0.25);
        }
        .hero {
                display:flex;
                justify-content:space-between;
                gap:18px;
                align-items:flex-start;
        }
        .hero h1 { margin: 0 0 4px; font-size: 28px; color: var(--text-main); }
        .hero-eyebrow {
                text-transform:uppercase;
                font-size:12px;
                letter-spacing:0.15em;
                color:var(--accent);
                margin:0 0 6px;
        }
        .meta { color: var(--text-muted); margin-top: 4px; font-size: 13px; }
        .hero-actions { display:flex; align-items:flex-start; }
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
        .metric-value--small { font-size: 20px; }
        .metric-caption { display:block; font-size:11px; color:var(--text-muted); margin-top:4px; }
        .metric--accent { border-color: rgba(37,99,235,0.35); background: #eef2ff; }
        .card-section { background: var(--card-bg); border: 1px solid var(--border); border-radius: 18px; padding: 22px 24px; margin-bottom: 26px; page-break-inside: avoid; }
        .card-section h2 { margin-top: 0; color: var(--accent); font-size: 19px; }
        .status-note { font-size:12px; color:var(--text-muted); margin-bottom:20px; }
        table { width:100%; border-collapse: collapse; font-size:11px; }
        thead { display: table-header-group; }
        tfoot { display: table-footer-group; }
        th, td { padding:8px 9px; border-bottom:1px solid var(--border); text-align:left; vertical-align:top; page-break-inside: avoid; }
        th { background:#eef2ff; font-weight:600; }
        tbody tr:nth-child(even) { background:#f9fafb; }
        tbody tr.row-keep { background: rgba(16, 185, 129, 0.15) !important; }
        tbody tr.row-delete { background: rgba(248, 113, 113, 0.18) !important; }
        tbody tr.row-review { background: rgba(250, 204, 21, 0.18) !important; }
        .table-wrapper { overflow:hidden; border-radius: 14px; border:1px solid var(--border); background:#fff; }
        .note { font-size:12px; color: var(--text-muted); margin-bottom: 18px; }
        .search-bar { margin: 0 0 14px; }
        .search-actions { display:flex; flex-wrap:wrap; gap:10px; align-items:center; }
        .search-input {
                width: 100%;
                max-width: 360px;
                padding: 8px 14px;
                border: 1px solid var(--border);
                border-radius: 999px;
                font-size: 13px;
                transition: border-color 0.2s, box-shadow 0.2s;
        }
        .search-input:focus {
                border-color: var(--accent);
                box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.15);
                outline: none;
        }
        .filter-controls {
                display:flex;
                flex-wrap:wrap;
                gap:10px;
                align-items:center;
                margin-top:8px;
        }
        .bulk-actions {
                display:flex;
                flex-wrap:wrap;
                gap:10px;
                align-items:center;
                font-size:12px;
                margin-top:8px;
        }
        .bulk-action-select {
                padding:6px 10px;
                border-radius:999px;
                border:1px solid var(--border);
                font-size:12px;
        }
        .bulk-apply-btn {
                padding:7px 14px;
                border-radius:999px;
                border:1px solid var(--accent);
                background:var(--accent);
                color:#fff;
                font-size:12px;
                cursor:pointer;
                transition:opacity 0.2s;
        }
        .bulk-apply-btn:disabled {
                opacity:0.5;
                cursor:not-allowed;
        }
        .filter-chips {
                display:flex;
                flex-wrap:wrap;
                gap:6px;
        }
        .chip {
                border-radius:999px;
                border:1px solid var(--border);
                padding:4px 12px;
                font-size:12px;
                background:#fff;
                cursor:pointer;
                transition:background 0.2s, color 0.2s, border-color 0.2s;
        }
        .chip.active {
                border-color:var(--accent);
                background:var(--accent);
                color:#fff;
        }
        .group-btn {
                padding: 8px 16px;
                border-radius: 999px;
                border: 1px solid var(--accent);
                background: #fff;
                color: var(--accent);
                font-size: 12px;
                cursor: pointer;
                transition: background 0.2s, color 0.2s;
        }
        .group-btn.active {
                background: var(--accent);
                color: #fff;
        }
        .admin-toggle {
                display:inline-flex;
                align-items:center;
                gap:6px;
                font-size:12px;
                color: var(--text-main);
        }
        .admin-toggle input {
                accent-color: var(--accent);
        }
        .export-btn {
                padding: 8px 16px;
                border-radius: 999px;
                border: 1px solid var(--accent);
                background: var(--accent);
                color: #fff;
                font-size: 12px;
                cursor: pointer;
                display: inline-flex;
                align-items: center;
                gap: 6px;
                transition: opacity 0.2s;
        }
        .export-btn:hover { opacity: 0.85; }
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
        .badge.review { background:#fef3c7; color:#92400e; }
        .bulk-select-cell, th.select-col {
                width:34px;
                text-align:center;
        }
        .bulk-select-cell input, th.select-col input {
                width:16px;
                height:16px;
                cursor:pointer;
        }
        .admin-icons { display:flex; flex-wrap:wrap; gap:4px; }
        .admin-icon {
                width:26px;
                height:26px;
                border-radius:999px;
                border:1px solid var(--border);
                background:#fff7ed;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:14px;
                line-height:1;
        }
        .admin-icon:hover { background:#fde68a; }
        .chart-section__head { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; gap:18px; }
        .chart-section__subtitle { margin:4px 0 0; font-size:13px; color:var(--text-muted); }
        .chart-grid {
                display:grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap:18px;
        }
        .chart-card {
                position:relative;
                background:linear-gradient(135deg, #eef2ff, #ffffff);
                border:1px solid rgba(37, 99, 235, 0.15);
                border-radius:20px;
                padding:20px 22px;
                box-shadow:0 25px 55px -30px rgba(15, 23, 42, 0.5);
                min-height:320px;
                display:flex;
                flex-direction:column;
                gap:16px;
                page-break-inside: avoid;
        }
        .chart-card--admin { background:linear-gradient(135deg, #f3e8ff, #ffffff); border-color: rgba(168, 85, 247, 0.35); }
        .chart-card--sku { background:linear-gradient(135deg, #ecfeff, #ffffff); border-color: rgba(45, 212, 191, 0.35); }
        .chart-eyebrow {
                text-transform:uppercase;
                letter-spacing:0.08em;
                font-size:11px;
                color:var(--text-muted);
                margin:0 0 4px;
        }
        .chart-card__head {
                display:flex;
                justify-content:space-between;
                gap:14px;
                align-items:flex-start;
        }
        .chart-card__head h3 {
                margin:0;
                font-size:18px;
                color:var(--text-main);
        }
        .chart-chip-tray {
                display:flex;
                flex-wrap:wrap;
                gap:8px;
                justify-content:flex-end;
        }
        .chart-chip {
                display:inline-flex;
                flex-direction:column;
        	footer { margin-top: 32px; color: var(--text-muted); font-size:11px; text-align:center; }
        	.small { font-size:11px; color: var(--text-muted); }
                .subtitle { font-size:11px; color: var(--text-muted); display:block; }
                border:1px solid rgba(15,23,42,0.1);
                background:rgba(255,255,255,0.7);
                font-size:11px;
                line-height:1.2;
        }
        .chart-chip strong {
                font-size:16px;
                color:var(--text-main);
                font-weight:600;
        }
        .chip-label {
                body.legacy-mode .hero { flex-direction:column; gap:8px; }
                font-size:11px;
                color:var(--text-muted);
        }
        .chart-chip--keep { border-color: rgba(16,185,129,0.3); }
        .chart-chip--review { border-color: rgba(250,204,21,0.35); }
        .chart-chip--delete { border-color: rgba(248,113,113,0.4); }
        .chart-chip--admin { border-color: rgba(168,85,247,0.4); }
        .chart-chip--neutral { border-color: rgba(148,163,184,0.45); }
        .chart-card__figure {
                position:relative;
                min-height:240px;
                border-radius:18px;
                background:rgba(255,255,255,0.85);
                padding:12px;
                overflow:hidden;
        }
        .chart-canvas {
                position:absolute;
                inset:12px;
                width:calc(100% - 24px);
                height:calc(100% - 24px);
        }
        .chart-center {
                position:absolute;
                top:50%;
                left:50%;
                transform:translate(-50%, -50%);
                text-align:center;
        }
        .chart-center__value {
                font-size:32px;
                font-weight:600;
                color:var(--accent);
                word-break:normal;
                white-space:nowrap;
                overflow:hidden;
                text-overflow:ellipsis;
        }
        .chart-center__label {
                font-size:12px;
                color:var(--text-muted);
                display:block;
                margin-top:4px;
                white-space:normal;
                word-break:break-word;
        }
        .chart-empty {
                position:absolute;
                inset:12px;
                display:none;
                align-items:center;
                justify-content:center;
                text-align:center;
                padding:24px;
                font-size:13px;
                color:var(--text-muted);
                background:rgba(255,255,255,0.95);
                border-radius:16px;
        }
        .chart-empty.is-visible { display:flex; }
        .chart-footnote {
                font-size:11px;
                color:var(--text-muted);
                margin:0;
        }
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
        body.legacy-mode tbody tr.row-review { background:#fff8db !important; }
        body.legacy-mode tbody tr.row-delete { background:#fff1f2 !important; }
        body.legacy-mode .badge.keep { background:#b7f7c3; color:#14532d; }
        body.legacy-mode .badge.review { background:#fde68a; color:#92400e; }
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
                .chart-grid { grid-template-columns: 1fr; gap: 10px; }
                .chart-card {
                        background:#fff !important;
                        border-color:#d1d5db;
                        box-shadow:none;
                }
                .chart-card__figure {
                        background:#fff;
                        border:1px solid rgba(15,23,42,0.12);
                        min-height:220px;
                }
                h2 { page-break-after: avoid; }
                .toolbar { display:none; }
                .search-bar { display:none !important; }
        }
        tr { page-break-inside: avoid; }
        .card-section { page-break-inside: avoid; }
</style>
</head>
<body>
<div class="report">
<header class="hero">
        <div>
                <p class="hero-eyebrow">$tenantDomainEsc</p>
                <h1>$(Get-LocalizedText 'Microsoft 365 Licensed Users Report' 'Rapport över licensierade Microsoft 365-användare')</h1>
                <div class="meta">
                        $(Get-LocalizedText 'Tenant' 'Klient'): $tenantDomainEsc &bull;
                        $(Get-LocalizedText 'Generated' 'Genererad'): $nowEsc
                </div>
        </div>
        <div class="hero-actions">
                <button id="legacyPrintButton" class="print-btn" type="button">🖨️ $(Get-LocalizedText 'Print classic layout' 'Skriv ut klassiskt läge')</button>
        </div>
</header>

<section class="summary">
        <div class="metric"><strong>$(Get-LocalizedText 'Total users' 'Totalt antal användare')</strong><div class="metric-value">$totalUsers</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Licensed users' 'Licensierade användare')</strong><div class="metric-value">$totalLicensed</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Unlicensed users' 'Olicensierade användare')</strong><div class="metric-value">$totalUnlicensed</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Keep recommended' 'Föreslås behållas')</strong><div class="metric-value" id="metric-keep">$keepCount</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Needs review' 'Behöver granskning')</strong><div class="metric-value" id="metric-review">$reviewCount</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Delete candidates' 'Föreslås tas bort')</strong><div class="metric-value" id="metric-delete">$deleteCount</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Admin accounts' 'Administratörskonton')</strong><div class="metric-value">$adminCount</div></div>
        <div class="metric"><strong>$(Get-LocalizedText 'Never signed in' 'Aldrig loggat in')</strong><div class="metric-value">$neverSignedInCount</div></div>
        <div class="metric metric--accent">
                <strong>$(Get-LocalizedText 'Next renewal' 'Nästa förnyelse')</strong>
                <div class="metric-value metric-value--small">$nextExpiryEsc</div>
                <span class="metric-caption">$(Get-LocalizedText 'Earliest tracked term end' 'Tidigt registrerat slutdatum')</span>
        </div>
</section>

<div class="note status-note">
        $(Get-LocalizedText "Keep = last login within $InactiveThresholdDays days (based on available sign-in logs). Review = approaching inactivity but still under evaluation. Delete = never signed in or past the threshold. Manual CSV overrides always win." "Behåll = senaste inloggning inom $InactiveThresholdDays dagar (baserat på tillgängliga inloggningsloggar). Granska = närmar sig inaktivitet men behöver bedömning. Ta bort = aldrig inloggad eller passerat tröskeln. Manuella CSV-överstyrningar gäller alltid.")
</div>

<section class="card-section chart-section">
        <div class="chart-section__head">
                <div>
                        <h2>$visualInsightsTitleEsc</h2>
                        <p class="chart-section__subtitle">$visualInsightsSubtitleEsc</p>
                </div>
        </div>
        <div class="chart-grid">
                <article class="chart-card chart-card--lifecycle">
                        <div class="chart-card__head">
                                <div>
                                        <p class="chart-eyebrow">$usagePulseLabelEsc</p>
                                        <h3>$lifecycleChartTitleEsc</h3>
                                </div>
                                <div class="chart-chip-tray">
                                        <div class="chart-chip chart-chip--keep">
                                                <strong id="chip-keep-percent">$lifecyclePercentKeepRounded%</strong>
                                                <span class="chip-label">$keepFilterLabelEsc</span>
                                        </div>
                                        <div class="chart-chip chart-chip--review">
                                                <strong id="chip-review-count">$reviewCount</strong>
                                                <span class="chip-label">$reviewFilterLabelEsc</span>
                                        </div>
                                        <div class="chart-chip chart-chip--delete">
                                                <strong id="chip-delete-count">$deleteCount</strong>
                                                <span class="chip-label">$deleteFilterLabelEsc</span>
                                        </div>
                                </div>
                        </div>
                        <div class="chart-card__figure">
                                <canvas id="lifecycleChart" class="chart-canvas" aria-label="$lifecycleChartTitleEsc"></canvas>
                                <div class="chart-center">
                                        <span class="chart-center__value" id="lifecycle-center-value">$keepCount</span>
                                        <span class="chart-center__label" id="lifecycle-center-label">$keepFilterLabelEsc</span>
                                </div>
                        </div>
                        <p class="chart-footnote">$chartLiveNoteEsc</p>
                </article>
                <article class="chart-card chart-card--admin">
                        <div class="chart-card__head">
                                <div>
                                        <p class="chart-eyebrow">$privExposureLabelEsc</p>
                                        <h3>$adminChartTitleEsc</h3>
                                </div>
                                <div class="chart-chip-tray">
                                        <div class="chart-chip chart-chip--admin">
                                                <strong id="chip-admin-percent">$adminPercentRounded%</strong>
                                                <span class="chip-label">$adminsShortLabelEsc</span>
                                        </div>
                                        <div class="chart-chip chart-chip--neutral">
                                                <strong id="chip-nonadmin-percent">$nonAdminPercentRounded%</strong>
                                                <span class="chip-label">$nonAdminsShortLabelEsc</span>
                                        </div>
                                </div>
                        </div>
                        <div class="chart-card__figure">
                                <canvas id="adminChart" class="chart-canvas" aria-label="$adminChartTitleEsc"></canvas>
                                <div class="chart-center">
                                        <span class="chart-center__value" id="admin-center-value">$adminCount</span>
                                        <span class="chart-center__label" id="admin-center-label">$adminsShortLabelEsc</span>
                                </div>
                        </div>
                        <p class="chart-footnote">$adminChartNoteEsc</p>
                </article>
                <article class="chart-card chart-card--sku">
                        <div class="chart-card__head">
                                <div>
                                        <p class="chart-eyebrow">$licenseMomentumLabelEsc</p>
                                        <h3>$skuChartTitleEsc</h3>
                                </div>
                                <div class="chart-chip-tray">
                                        <div class="chart-chip chart-chip--neutral">
                                                <strong id="chip-topsku-share">$topSkuLeaderPercentRounded%</strong>
                                                <span class="chip-label">$shareLabelEsc</span>
                                        </div>
                                        <div class="chart-chip chart-chip--neutral">
                                                <strong id="chip-topsku-seats">$topSkuLeaderCount</strong>
                                                <span class="chip-label">$seatsLabelEsc</span>
                                        </div>
                                </div>
                        </div>
                        <div class="chart-card__figure">
                                <canvas id="skuChart" class="chart-canvas" aria-label="$skuChartTitleEsc"></canvas>
                                <div class="chart-center">
                                        <span class="chart-center__value" id="sku-center-value">$topSkuCenterValueEsc</span>
                                        <span class="chart-center__label" id="sku-center-label"$skuCenterLabelTitleAttr>$topSkuCenterLabelEsc</span>
                                </div>
                                <div class="chart-empty" data-empty-text="$noSkuDataLabelAttr"></div>
                        </div>
                        <p class="chart-footnote">$skuChartNoteEsc</p>
                </article>
        </div>
</section>
"@

if ($skuPurchaseTable.Count -gt 0) {
$reportHtml += @"
<section class="card-section">
        <h2>$(Get-LocalizedText 'Subscription purchases' 'Prenumerationsköp')</h2>
    <div class="table-wrapper">
        <table class="subscription-table">
                <thead>
                        <tr>
                                <th>$(Get-LocalizedText 'Product' 'Produkt')</th>
                                <th>$(Get-LocalizedText 'SKU' 'SKU')</th>
                                <th>$(Get-LocalizedText 'Term start' 'Startdatum')</th>
                                <th>$(Get-LocalizedText 'Term end' 'Slutdatum')</th>
                                <th>$(Get-LocalizedText 'Days remaining' 'Dagar kvar')</th>
                                <th>$(Get-LocalizedText 'Purchased' 'Köpta')</th>
                                <th>$(Get-LocalizedText 'Consumed' 'Förbrukade')</th>
                        </tr>
                </thead>
        <tbody>
"@

        foreach ($row in $skuPurchaseTable) {
                $product = [System.Web.HttpUtility]::HtmlEncode($row.FriendlyName)
                if ($row.IsTrial) {
                        $product += " <span class='badge delete'>$trialLabel</span>"
                }
                $skuPart = [System.Web.HttpUtility]::HtmlEncode($row.SkuPartNumber)
                $termStartDisplay = if ($row.TermStart) { [System.Web.HttpUtility]::HtmlEncode($row.TermStart.ToString('yyyy-MM-dd')) } else { $unknownLabel }
                $termEndDisplay = if ($row.TermEnd) { [System.Web.HttpUtility]::HtmlEncode($row.TermEnd.ToString('yyyy-MM-dd')) } else { $unknownLabel }
                $daysDisplay = if ($null -ne $row.DaysRemaining) { [string]$row.DaysRemaining } else { '—' }
                $purchasedDisplay = if ($null -ne $row.PurchasedUnits) { [string]$row.PurchasedUnits } else { $naLabel }
                $consumedDisplay = if ($null -ne $row.ConsumedUnits) { [string]$row.ConsumedUnits } else { $naLabel }
                $reportHtml += "        <tr><td>$product</td><td>$skuPart</td><td>$termStartDisplay</td><td>$termEndDisplay</td><td>$daysDisplay</td><td>$purchasedDisplay</td><td>$consumedDisplay</td></tr>`n"
        }

        $reportHtml += @"
        </tbody>
        </table>
    </div>
        <div class="note small">$([System.Web.HttpUtility]::HtmlEncode($subscriptionTableNote))</div>
</section>
"@
}
else {
        $reportHtml += @"
<section class="card-section">
        <h2>$(Get-LocalizedText 'Subscription purchases' 'Prenumerationsköp')</h2>
                <div class="note">$([System.Web.HttpUtility]::HtmlEncode($subscriptionMetadataFallbackMessage))</div>
</section>
"@
}

$reportHtml += @"
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
        <div class="search-bar">
                <div class="search-actions">
                        <input type="search" id="user-search" class="search-input" placeholder="$searchPlaceholderEsc" aria-label="$searchPlaceholderEsc" autocomplete="off" spellcheck="false">
                        <button type="button" id="group-toggle" class="group-btn" data-group-label="$groupByLabelEsc" data-ungroup-label="$groupResetLabelEsc">$groupByLabelEsc</button>
                </div>
                <div class="filter-controls">
                        <div class="filter-chips" role="toolbar" aria-label="$statusFilterToolbarLabelEsc">
                                <button type="button" class="chip filter-chip active" data-status-filter="all">$statusFilterAllLabelEsc</button>
                                <button type="button" class="chip filter-chip" data-status-filter="keep">$keepFilterLabelEsc</button>
                                <button type="button" class="chip filter-chip" data-status-filter="review">$reviewFilterLabelEsc</button>
                                <button type="button" class="chip filter-chip" data-status-filter="delete">$deleteFilterLabelEsc</button>
                        </div>
                        <label class="admin-toggle">
                                <input type="checkbox" id="admins-only-toggle">
                                <span>$showAdminsOnlyLabelEsc</span>
                        </label>
                        <button type="button" id="export-csv" class="export-btn" data-export-filename="$csvFileNameEsc" data-empty-message="$exportEmptyMessageEsc">⬇ $exportCsvLabelEsc</button>
                </div>
                <div class="bulk-actions" role="group" aria-label="$(Get-LocalizedText 'Bulk edit actions' 'Massuppdateringar')">
                        <span>$(Get-LocalizedText 'Bulk edit' 'Massändring')</span>
                        <select id="bulk-action-select" class="bulk-action-select">
                                <option value="">$(Get-LocalizedText 'Choose action' 'Välj åtgärd')</option>
                                <option value="keep">$(Get-LocalizedText 'Mark as Keep' 'Markera som Behåll')</option>
                                <option value="review">$(Get-LocalizedText 'Mark as Review' 'Markera som Granska')</option>
                                <option value="delete">$(Get-LocalizedText 'Mark as Delete' 'Markera som Ta bort')</option>
                        </select>
                        <button type="button" id="apply-bulk-action" class="bulk-apply-btn" disabled>$(Get-LocalizedText 'Apply' 'Utför') (<span id="bulk-selection-count">0</span>)</button>
                </div>
        </div>
    <div class="table-wrapper">
        <table id="licensed-users-table">
                <thead>
                        <tr>
                                <th class="select-col">
                                        <input type="checkbox" id="select-all-rows" aria-label="$(Get-LocalizedText 'Select all users' 'Markera alla användare')">
                                </th>
                                <th>$(Get-LocalizedText 'Name' 'Namn')</th>
                                <th>$(Get-LocalizedText 'UPN' 'UPN')</th>
                                <th>$(Get-LocalizedText 'Email' 'E-post')</th>
                                <th>$(Get-LocalizedText 'Last login' 'Senaste inloggning')</th>
                                <th>$(Get-LocalizedText 'Action' 'Åtgärd')</th>
                                <th>$(Get-LocalizedText 'Admin rights' 'Administratörsrättigheter')</th>
                                <th>$(Get-LocalizedText 'Licenses' 'Licenser')</th>
                                <th>$(Get-LocalizedText 'Count' 'Antal')</th>
                        </tr>
                </thead>
        <tbody>
"@

$selectUserLabelText = Get-LocalizedText 'Select user' 'Markera användare'
foreach ($u in $licensedUsers) {
        $name = [System.Web.HttpUtility]::HtmlEncode($u.DisplayName)
        $upn = [System.Web.HttpUtility]::HtmlEncode($u.UserPrincipalName)
        $mail = [System.Web.HttpUtility]::HtmlEncode($u.Mail)
        $lics = [System.Web.HttpUtility]::HtmlEncode($u.Licenses)
        $licensesAttr = if ($u.Licenses) { [System.Web.HttpUtility]::HtmlAttributeEncode($u.Licenses) } else { '' }
        $lastLogin = [System.Web.HttpUtility]::HtmlEncode($u.LastLoginDisplay)
        $actionLabel = [System.Web.HttpUtility]::HtmlEncode($u.LifecycleStatusLabel)
        $rowClassAttr = switch ($u.LifecycleStatus) {
                'keep' { 'row-keep' }
                'review' { 'row-review' }
                default { 'row-delete' }
        }
        $badgeClassAttr = switch ($u.LifecycleStatus) {
                'keep' { 'badge keep' }
                'review' { 'badge review' }
                default { 'badge delete' }
        }
        $keepLabelEsc = [System.Web.HttpUtility]::HtmlEncode($keepLabel)
        $deleteLabelEsc = [System.Web.HttpUtility]::HtmlEncode($deleteLabel)
        $reviewLabelEsc = [System.Web.HttpUtility]::HtmlEncode((Get-LocalizedText 'Review' 'Granska'))
        $statusValue = [System.Web.HttpUtility]::HtmlEncode($u.LifecycleStatus)
        $userIdentifier = if ($u.UserPrincipalName) { $u.UserPrincipalName } elseif ($u.Mail) { $u.Mail } else { $u.DisplayName }
        $userIdentifierEsc = [System.Web.HttpUtility]::HtmlEncode($userIdentifier)
        $cnt = $u.LicenseCount
        $adminIconsHtml = '—'
        if ($u.AdminRoles -and $u.AdminRoles.Count -gt 0) {
                $iconFragments = foreach ($roleInfo in ($u.AdminRoles | Sort-Object Name)) {
                        $iconValue = if ($roleInfo.Icon) { $roleInfo.Icon } else { '🛡️' }
                        $iconSymbol = [System.Web.HttpUtility]::HtmlEncode($iconValue)
                        $roleNameEsc = [System.Web.HttpUtility]::HtmlEncode($roleInfo.Name)
                        "<span class='admin-icon' title='$roleNameEsc'>$iconSymbol</span>"
                }
                $adminIconsHtml = "<div class='admin-icons'>" + ($iconFragments -join '') + "</div>"
        }
        $hasAdminFlag = if ($u.AdminRoles -and $u.AdminRoles.Count -gt 0) { '1' } else { '0' }
        $searchTerms = (@($u.DisplayName, $u.UserPrincipalName, $u.Mail, $u.Licenses) | Where-Object { $_ })
        $searchBlobSource = ($searchTerms -join ' ').Trim()
        $searchBlob = if ([string]::IsNullOrWhiteSpace($searchBlobSource)) { '' } else { $searchBlobSource.ToLowerInvariant() }
        $searchBlobEsc = [System.Web.HttpUtility]::HtmlEncode($searchBlob)
        $licenseKey = if ([string]::IsNullOrWhiteSpace($u.Licenses)) { '' } else { $u.Licenses.ToLowerInvariant() }
        $licenseKeyEsc = [System.Web.HttpUtility]::HtmlEncode($licenseKey)
        $nameKey = if ([string]::IsNullOrWhiteSpace($u.DisplayName)) { '' } else { $u.DisplayName.ToLowerInvariant() }
        $nameKeyEsc = [System.Web.HttpUtility]::HtmlEncode($nameKey)
        $selectionTarget = if (-not [string]::IsNullOrWhiteSpace($u.DisplayName)) { $u.DisplayName } elseif (-not [string]::IsNullOrWhiteSpace($u.UserPrincipalName)) { $u.UserPrincipalName } elseif (-not [string]::IsNullOrWhiteSpace($u.Mail)) { $u.Mail } else { $userIdentifier }
        $selectionAriaText = ("$selectUserLabelText $selectionTarget").Trim()
        $selectionAriaAttr = [System.Web.HttpUtility]::HtmlAttributeEncode($selectionAriaText)
        $reportHtml += "    <tr class='$rowClassAttr' data-status='$statusValue' data-keep-label='$keepLabelEsc' data-delete-label='$deleteLabelEsc' data-review-label='$reviewLabelEsc' data-user-id='$userIdentifierEsc' data-search='$searchBlobEsc' data-search-visible='1' data-license-key='$licenseKeyEsc' data-name-key='$nameKeyEsc' data-has-admin='$hasAdminFlag' data-licenses='$licensesAttr'><td class='bulk-select-cell'><input type='checkbox' class='bulk-select' aria-label='$selectionAriaAttr'></td><td>$name</td><td>$upn</td><td>$mail</td><td>$lastLogin</td><td><span class='$badgeClassAttr status-toggle'>$actionLabel</span></td><td>$adminIconsHtml</td><td>$lics</td><td>$cnt</td></tr>`n"
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

        // Cache the high-touch DOM nodes that power filters, KPIs, and canvas charts so large tenants stay snappy.
        const keepMetric = document.getElementById('metric-keep');
        const deleteMetric = document.getElementById('metric-delete');
        const reviewMetric = document.getElementById('metric-review');
        const statusOrder = ['keep','review','delete'];
        const searchInput = document.getElementById('user-search');
        const usersTable = document.getElementById('licensed-users-table');
        const usersTbody = usersTable ? usersTable.querySelector('tbody') : null;
        const userRows = usersTbody ? Array.from(usersTbody.querySelectorAll('tr[data-status]')) : [];
        const groupToggleBtn = document.getElementById('group-toggle');
        const filterChipButtons = Array.from(document.querySelectorAll('.filter-chip'));
        const adminsOnlyToggle = document.getElementById('admins-only-toggle');
        const exportButton = document.getElementById('export-csv');
        const selectAllCheckbox = document.getElementById('select-all-rows');
        const bulkActionSelect = document.getElementById('bulk-action-select');
        const bulkApplyButton = document.getElementById('apply-bulk-action');
        const bulkSelectionCount = document.getElementById('bulk-selection-count');
        const lifecycleCenterValue = document.getElementById('lifecycle-center-value');
        const lifecycleCenterLabel = document.getElementById('lifecycle-center-label');
        const adminCenterValue = document.getElementById('admin-center-value');
        const adminCenterLabel = document.getElementById('admin-center-label');
        const skuCenterValue = document.getElementById('sku-center-value');
        const skuCenterLabel = document.getElementById('sku-center-label');
        const chipKeepPercent = document.getElementById('chip-keep-percent');
        const chipReviewCount = document.getElementById('chip-review-count');
        const chipDeleteCount = document.getElementById('chip-delete-count');
        const chipAdminPercent = document.getElementById('chip-admin-percent');
        const chipNonAdminPercent = document.getElementById('chip-nonadmin-percent');
        const chipTopSkuShare = document.getElementById('chip-topsku-share');
        const chipTopSkuSeats = document.getElementById('chip-topsku-seats');
        const lifecycleCanvas = document.getElementById('lifecycleChart');
        const adminCanvas = document.getElementById('adminChart');
        const skuCanvas = document.getElementById('skuChart');
        const skuEmptyState = document.querySelector('.chart-card--sku .chart-empty');
        const originalOrder = userRows.slice();
        let grouped = false;
        let activeStatusFilter = 'all';
        let currentSearchTerm = '';
        let adminsOnly = false;
        let lastSummary = null;
        let lastSkuSeries = null;
        let resizeTimer = null;
        let resizeObserver = null;
        let windowResizeHooked = false;

        function truncateLabel(label, maxLen){
                if (!label) { return ''; }
                const trimmed = String(label).trim();
                if (trimmed.length <= (maxLen || 32)) { return trimmed; }
                const limit = Math.max((maxLen || 32) - 1, 1);
                return trimmed.slice(0, limit) + '…';
        }

        function withScrollLock(work){
                if (typeof work !== 'function') { return; }
                const scrollX = window.scrollX || document.documentElement.scrollLeft || 0;
                const scrollY = window.scrollY || document.documentElement.scrollTop || 0;
                // Many bulk actions re-render the table; lock the scroll position so the operator keeps context.
                work();
                window.requestAnimationFrame(function(){
                        window.scrollTo(scrollX, scrollY);
                });
        }

        function normalizeStatus(value) {
                if (!value) { return 'review'; }
                const lowered = value.toLowerCase();
                return statusOrder.includes(lowered) ? lowered : 'review';
        }

        function applyStatus(row, status) {
                if(!row) { return; }
                const normalized = normalizeStatus(status);
                row.dataset.status = normalized;
                row.classList.remove('row-keep','row-review','row-delete');
                row.classList.add('row-' + normalized);
                const badge = row.querySelector('.status-toggle');
                if (badge) {
                        badge.classList.remove('keep','review','delete');
                        badge.classList.add(normalized);
                        const keepLabel = row.dataset.keepLabel || 'Keep';
                        const deleteLabel = row.dataset.deleteLabel || 'Delete';
                        const reviewLabel = row.dataset.reviewLabel || 'Review';
                        const labelMap = { keep: keepLabel, review: reviewLabel, delete: deleteLabel };
                        badge.textContent = labelMap[normalized] || normalized;
                }
        }

        function prepareCanvas(canvas){
                if (!canvas) { return null; }
                const rect = canvas.getBoundingClientRect();
                const width = Math.max(rect.width || 280, 200);
                const height = Math.max(rect.height || 230, 200);
                const ratio = window.devicePixelRatio || 1;
                // Resize the backing store on every render so exported PDFs stay crisp on Retina/high-DPI screens.
                canvas.width = width * ratio;
                canvas.height = height * ratio;
                const ctx = canvas.getContext('2d');
                if (!ctx) { return null; }
                ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
                ctx.clearRect(0, 0, width, height);
                return { ctx, width, height };
        }

        function drawDonutChart(canvas, segments){
                const meta = prepareCanvas(canvas);
                if (!meta) { return; }
                const { ctx, width, height } = meta;
                const size = Math.min(width, height);
                const radius = Math.max(Math.min(size / 2 - 12, size * 0.45), 36);
                const thickness = Math.max(radius * 0.42, 18);
                const adjustedRadius = radius - thickness / 2;
                ctx.lineWidth = thickness;
                ctx.lineCap = 'butt';
                ctx.strokeStyle = 'rgba(226,232,240,0.85)';
                ctx.beginPath();
                ctx.arc(width / 2, height / 2, adjustedRadius, 0, Math.PI * 2);
                ctx.stroke();
                const total = segments.reduce(function(sum, segment){
                        return sum + Math.max(segment.value || 0, 0);
                }, 0);
                if (total <= 0) { return; }
                let startAngle = -Math.PI / 2;
                segments.forEach(function(segment){
                        const value = Math.max(segment.value || 0, 0);
                        if (value <= 0) { return; }
                        const sweep = (value / total) * Math.PI * 2;
                        ctx.beginPath();
                        ctx.strokeStyle = segment.color;
                        ctx.arc(width / 2, height / 2, adjustedRadius, startAngle, startAngle + sweep);
                        ctx.stroke();
                        startAngle += sweep;
                });
        }

        function drawHorizontalBars(canvas, bars){
                const meta = prepareCanvas(canvas);
                if (!meta) { return; }
                const { ctx, width, height } = meta;
                const paddingX = 18;
                const verticalPadding = 28;
                const count = bars.length;
                const availableHeight = Math.max(height - verticalPadding * 2, 40);
                const gap = count > 1 ? Math.min(12, availableHeight / (count * 2)) : 0;
                const rowHeight = count > 0 ? Math.max((availableHeight - gap * (count - 1)) / count, 14) : 18;
                const labelArea = Math.min(200, width * 0.45);
                const barStart = paddingX + labelArea + 12;
                const barWidth = Math.max(width - barStart - paddingX, 40);
                const maxValue = bars.reduce(function(max, bar){
                        return Math.max(max, bar.value || 0);
                }, 0) || 1;
                ctx.font = '12px "Segoe UI", "Helvetica Neue", Arial, sans-serif';
                ctx.textBaseline = 'middle';
                bars.forEach(function(bar, index){
                        const centerY = verticalPadding + index * (rowHeight + gap) + rowHeight / 2;
                        ctx.fillStyle = '#0f172a';
                        ctx.textAlign = 'left';
                        ctx.fillText(truncateLabel(bar.label || '', 50), paddingX, centerY);
                        ctx.fillStyle = 'rgba(203,213,225,0.6)';
                        ctx.fillRect(barStart, centerY - rowHeight / 3, barWidth, (rowHeight / 3) * 2);
                        const fillWidth = Math.max(0, Math.min(barWidth, (bar.value / maxValue) * barWidth));
                        ctx.fillStyle = '#0ea5e9';
                        ctx.fillRect(barStart, centerY - rowHeight / 3, fillWidth, (rowHeight / 3) * 2);
                        ctx.fillStyle = '#0f172a';
                        ctx.textAlign = 'right';
                        ctx.fillText(String(bar.value || 0), barStart + barWidth, centerY);
                });
                if (count === 0) {
                        ctx.fillStyle = '#94a3b8';
                        ctx.textAlign = 'center';
                        ctx.fillText('—', width / 2, height / 2);
                }
        }

        function renderCharts(){
                if (!lastSummary) { return; }
                if (lifecycleCanvas) {
                        drawDonutChart(lifecycleCanvas, [
                                { value: lastSummary.keepCount, color: '#16a34a' },
                                { value: lastSummary.reviewCount, color: '#facc15' },
                                { value: lastSummary.deleteCount, color: '#ef4444' }
                        ]);
                }
                if (adminCanvas) {
                        drawDonutChart(adminCanvas, [
                                { value: lastSummary.adminVisible, color: '#7c3aed' },
                                { value: Math.max(lastSummary.totalVisible - lastSummary.adminVisible, 0), color: '#cbd5f5' }
                        ]);
                }
                if (skuCanvas && lastSkuSeries) {
                        const bars = lastSkuSeries.labels.map(function(label, idx){
                                return { label: label, value: lastSkuSeries.values[idx] };
                        });
                        drawHorizontalBars(skuCanvas, bars);
                }
        }

        function scheduleChartRedraw(){
                if (!lastSummary) { return; }
                clearTimeout(resizeTimer);
                resizeTimer = window.setTimeout(renderCharts, 80);
        }

        function watchChartResizing(){
                const targets = [lifecycleCanvas, adminCanvas, skuCanvas].map(function(canvas){
                        if (!canvas) { return null; }
                        return canvas.closest('.chart-card__figure') || canvas.parentElement || canvas;
                }).filter(Boolean);

                if (typeof ResizeObserver !== 'undefined') {
                        if (resizeObserver) {
                                resizeObserver.disconnect();
                        }
                        resizeObserver = new ResizeObserver(function(entries){
                                if (!entries || entries.length === 0) { return; }
                                scheduleChartRedraw();
                        });
                        targets.forEach(function(target){ resizeObserver.observe(target); });
                }
                else if (!windowResizeHooked) {
                        window.addEventListener('resize', scheduleChartRedraw);
                        windowResizeHooked = true;
                }
        }

        watchChartResizing();

        function buildSkuSeries(visibleRows, totalVisible){
                const counts = new Map();
                (visibleRows || []).forEach(function(row){
                        const licensesRaw = row.dataset.licenses || '';
                        if (!licensesRaw) { return; }
                        const tokens = licensesRaw.split(/,\s*/).map(function(part){ return part.trim(); }).filter(Boolean);
                        const uniqueTokens = Array.from(new Set(tokens));
                        uniqueTokens.forEach(function(label){
                                const nextValue = (counts.get(label) || 0) + 1;
                                counts.set(label, nextValue);
                        });
                });
                const sorted = Array.from(counts.entries()).sort(function(a, b){ return b[1] - a[1]; }).slice(0, 5);
                const leader = sorted[0];
                const leaderCount = leader ? leader[1] : 0;
                const leaderName = leader ? leader[0] : '';
                const leaderShare = (leaderCount > 0 && totalVisible > 0) ? Math.round((leaderCount / totalVisible) * 100) : 0;
                return {
                        labels: sorted.map(function(entry){ return entry[0]; }),
                        values: sorted.map(function(entry){ return entry[1]; }),
                        leaderName: leaderName,
                        leaderCount: leaderCount,
                        leaderShare: leaderShare
                };
        }

        function updateChartMetrics(summary){
                if (!summary) { return; }
                const lifecycleTotals = [summary.keepCount, summary.reviewCount, summary.deleteCount];
                const lifecycleTotal = lifecycleTotals.reduce(function(sum, value){ return sum + value; }, 0);
                const keepPercentDisplay = lifecycleTotal ? Math.round((summary.keepCount / lifecycleTotal) * 100) : 0;
                const adminTotals = [summary.adminVisible, Math.max(summary.totalVisible - summary.adminVisible, 0)];
                const adminPercentDisplay = summary.totalVisible ? Math.round((summary.adminVisible / summary.totalVisible) * 100) : 0;
                const nonAdminPercentDisplay = summary.totalVisible ? Math.round((adminTotals[1] / summary.totalVisible) * 100) : 0;
                if (keepMetric) { keepMetric.textContent = summary.keepCount; }
                if (deleteMetric) { deleteMetric.textContent = summary.deleteCount; }
                if (reviewMetric) { reviewMetric.textContent = summary.reviewCount; }
                if (chipKeepPercent) { chipKeepPercent.textContent = keepPercentDisplay + '%'; }
                if (chipReviewCount) { chipReviewCount.textContent = summary.reviewCount; }
                if (chipDeleteCount) { chipDeleteCount.textContent = summary.deleteCount; }
                if (chipAdminPercent) { chipAdminPercent.textContent = adminPercentDisplay + '%'; }
                if (chipNonAdminPercent) { chipNonAdminPercent.textContent = nonAdminPercentDisplay + '%'; }
                if (lifecycleCenterValue) { lifecycleCenterValue.textContent = summary.keepCount; }
                if (adminCenterValue) { adminCenterValue.textContent = summary.adminVisible; }
                const skuSeries = buildSkuSeries(summary.visibleRows, summary.totalVisible);
                const hasSkuData = skuSeries.values.some(function(value){ return value > 0; });
                const leaderLabel = skuSeries.leaderName || '';
                const hasLeader = hasSkuData && Boolean(leaderLabel);
                const leaderShareDisplay = hasLeader ? (skuSeries.leaderShare > 0 ? skuSeries.leaderShare + '%' : String(skuSeries.leaderCount)) : '—';
                if (chipTopSkuShare) { chipTopSkuShare.textContent = hasSkuData ? (skuSeries.leaderShare + '%') : '0%'; }
                if (chipTopSkuSeats) { chipTopSkuSeats.textContent = hasSkuData ? skuSeries.leaderCount : '0'; }
                if (skuCenterValue) { skuCenterValue.textContent = leaderShareDisplay; }
                if (skuCenterLabel) {
                        if (hasLeader) {
                                const truncated = truncateLabel(leaderLabel, 36);
                                skuCenterLabel.textContent = truncated;
                                skuCenterLabel.title = leaderLabel;
                        } else if (skuEmptyState) {
                                skuCenterLabel.textContent = skuEmptyState.dataset.emptyText || '—';
                                skuCenterLabel.removeAttribute('title');
                        } else {
                                skuCenterLabel.textContent = '—';
                                skuCenterLabel.removeAttribute('title');
                        }
                }
                if (skuEmptyState) {
                        if (hasSkuData) {
                                skuEmptyState.classList.remove('is-visible');
                                skuEmptyState.textContent = '';
                        } else {
                                skuEmptyState.textContent = skuEmptyState.dataset.emptyText || 'Not enough license data yet';
                                skuEmptyState.classList.add('is-visible');
                        }
                }
                lastSummary = summary;
                lastSkuSeries = skuSeries;
                renderCharts();
        }

        function recalcSummary(){
                const visibleRows = [];
                let keepCount = 0;
                let deleteCount = 0;
                let reviewCount = 0;
                let adminVisible = 0;
                userRows.forEach(function(row){
                        if (row.dataset.searchVisible === '0') { return; }
                        visibleRows.push(row);
                        const status = normalizeStatus(row.dataset.status);
                        if (status === 'keep') { keepCount++; }
                        else if (status === 'review') { reviewCount++; }
                        else { deleteCount++; }
                        if (row.dataset.hasAdmin === '1') { adminVisible++; }
                });
                return {
                        keepCount: keepCount,
                        deleteCount: deleteCount,
                        reviewCount: reviewCount,
                        adminVisible: adminVisible,
                        totalVisible: visibleRows.length,
                        visibleRows: visibleRows
                };
        }

        function updateRowVisibility(row){
                const haystack = row.dataset.search || '';
                const statusValue = normalizeStatus(row.dataset.status);
                const matchesSearch = currentSearchTerm === '' || haystack.includes(currentSearchTerm);
                const matchesStatus = activeStatusFilter === 'all' || statusValue === activeStatusFilter;
                const matchesAdmin = !adminsOnly || row.dataset.hasAdmin === '1';
                const visible = matchesSearch && matchesStatus && matchesAdmin;
                row.dataset.searchVisible = visible ? '1' : '0';
                row.style.display = visible ? '' : 'none';
        }

        function applyFilters(){
                if (userRows.length === 0) { return; }
                withScrollLock(function(){
                        userRows.forEach(updateRowVisibility);
                });
                const summary = recalcSummary();
                updateChartMetrics(summary);
        }

        function getVisibleRows(){
                return userRows.filter(function(row){ return row.dataset.searchVisible !== '0'; });
        }

        function getRowCheckbox(row){
                if (!row) { return null; }
                return row.querySelector('.bulk-select');
        }

        function getSelectedRows(){
                return userRows.filter(function(row){
                        const checkbox = getRowCheckbox(row);
                        return checkbox && checkbox.checked;
                });
        }

        function refreshBulkUi(){
                const selectedRows = getSelectedRows();
                const selectedCount = selectedRows.length;
                if (bulkSelectionCount) { bulkSelectionCount.textContent = selectedCount; }
                if (bulkApplyButton) {
                        const actionValue = bulkActionSelect ? (bulkActionSelect.value || '') : '';
                        bulkApplyButton.disabled = !(selectedCount > 0 && actionValue);
                }
                if (selectAllCheckbox) {
                        const totalRows = userRows.length;
                        selectAllCheckbox.checked = totalRows > 0 && selectedCount === totalRows;
                        selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < totalRows;
                }
                return selectedRows;
        }

        function downloadFilteredCsv(){
                if (!exportButton || typeof Blob === 'undefined') { return; }
                const visibleRows = getVisibleRows();
                if (visibleRows.length === 0) {
                        const emptyMessage = exportButton.dataset.emptyMessage || 'No rows match the current filters.';
                        window.alert(emptyMessage);
                        return;
                }
                const header = ['Name','UPN','Email','Last login','Action','Admin rights','Licenses','License count'];
                const csvMatrix = [header];
                visibleRows.forEach(function(row){
                        const cells = row.querySelectorAll('td');
                        if (cells.length < 8) { return; }
                        const dataColumns = 8;
                        const startIndex = Math.max(0, cells.length - dataColumns);
                        const rowValues = [
                                cells[startIndex + 0].textContent.trim(),
                                cells[startIndex + 1].textContent.trim(),
                                cells[startIndex + 2].textContent.trim(),
                                cells[startIndex + 3].textContent.trim(),
                                cells[startIndex + 4].textContent.trim(),
                                cells[startIndex + 5].textContent.replace(/\s+/g, ' ').trim(),
                                cells[startIndex + 6].textContent.replace(/\s+/g, ' ').trim(),
                                cells[startIndex + 7].textContent.trim()
                        ];
                        csvMatrix.push(rowValues);
                });
                const csvString = csvMatrix.map(function(columns){
                        return columns.map(function(value){
                                const sanitized = value.replace(/"/g, '""');
                                return '"' + sanitized + '"';
                        }).join(',');
                }).join('\r\n');
                const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
                const urlCreator = window.URL || window.webkitURL;
                if (!urlCreator) { return; }
                const url = urlCreator.createObjectURL(blob);
                const anchor = document.createElement('a');
                anchor.href = url;
                anchor.download = exportButton.dataset.exportFilename || 'licensed-users.csv';
                document.body.appendChild(anchor);
                anchor.click();
                setTimeout(function(){
                        document.body.removeChild(anchor);
                        urlCreator.revokeObjectURL(url);
                }, 0);
        }

        userRows.forEach(function(row){
                const badge = row.querySelector('.status-toggle');
                if (badge) {
                        badge.addEventListener('click', function(){
                                const current = normalizeStatus(row.dataset.status);
                                const index = statusOrder.indexOf(current);
                                const nextStatus = statusOrder[(index + 1) % statusOrder.length];
                                applyStatus(row, nextStatus);
                                applyFilters();
                        });
                }
                const checkbox = row.querySelector('.bulk-select');
                if (checkbox) {
                        checkbox.addEventListener('change', refreshBulkUi);
                }
        });

        if (selectAllCheckbox) {
                selectAllCheckbox.addEventListener('change', function(){
                        const checked = selectAllCheckbox.checked;
                        userRows.forEach(function(row){
                                const checkbox = row.querySelector('.bulk-select');
                                if (checkbox) {
                                        checkbox.checked = checked;
                                }
                        });
                        refreshBulkUi();
                });
        }

        if (bulkActionSelect) {
                bulkActionSelect.addEventListener('change', refreshBulkUi);
        }

        if (bulkApplyButton) {
                bulkApplyButton.addEventListener('click', function(){
                        const rawAction = bulkActionSelect ? (bulkActionSelect.value || '') : '';
                        const actionValue = rawAction ? normalizeStatus(rawAction) : '';
                        const selectedRows = getSelectedRows();
                        if (!actionValue || selectedRows.length === 0) { return; }
                        selectedRows.forEach(function(row){
                                applyStatus(row, actionValue);
                        });
                        applyFilters();
                        refreshBulkUi();
                });
        }

        function groupByLicense(){
                if (!usersTbody) { return; }
                withScrollLock(function(){
                        const sorted = userRows.slice().sort(function(a, b){
                                const aKey = a.dataset.licenseKey || '';
                                const bKey = b.dataset.licenseKey || '';
                                if (aKey === bKey) {
                                        const aName = a.dataset.nameKey || '';
                                        const bName = b.dataset.nameKey || '';
                                        return aName.localeCompare(bName);
                                }
                                return aKey.localeCompare(bKey);
                        });
                        sorted.forEach(function(row){ usersTbody.appendChild(row); });
                });
        }

        function restoreOriginalOrder(){
                if (!usersTbody) { return; }
                withScrollLock(function(){
                        originalOrder.forEach(function(row){ usersTbody.appendChild(row); });
                });
        }

        if (groupToggleBtn && usersTbody) {
                groupToggleBtn.addEventListener('click', function(){
                        grouped = !grouped;
                        if (grouped) {
                                groupByLicense();
                                groupToggleBtn.classList.add('active');
                                groupToggleBtn.textContent = groupToggleBtn.dataset.ungroupLabel || 'Show original order';
                        } else {
                                restoreOriginalOrder();
                                groupToggleBtn.classList.remove('active');
                                groupToggleBtn.textContent = groupToggleBtn.dataset.groupLabel || 'Group by subscription type';
                        }
                });
        }

        if (searchInput) {
                const updateSearchTerm = function(){
                        currentSearchTerm = (searchInput.value || '').trim().toLowerCase();
                        applyFilters();
                };
                searchInput.addEventListener('input', updateSearchTerm);
                updateSearchTerm();
        } else {
                applyFilters();
        }

        if (filterChipButtons.length > 0) {
                filterChipButtons.forEach(function(chip){
                        chip.addEventListener('click', function(){
                                activeStatusFilter = chip.dataset.statusFilter || 'all';
                                filterChipButtons.forEach(function(btn){ btn.classList.toggle('active', btn === chip); });
                                applyFilters();
                        });
                });
        }

        if (adminsOnlyToggle) {
                adminsOnlyToggle.addEventListener('change', function(){
                        adminsOnly = adminsOnlyToggle.checked;
                        applyFilters();
                });
        }

        if (exportButton) {
                exportButton.addEventListener('click', function(){
                        downloadFilteredCsv();
                });
        }

        refreshBulkUi();
        watchChartResizing();
        applyFilters();
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