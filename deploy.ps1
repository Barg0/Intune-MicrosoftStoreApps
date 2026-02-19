# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script Name ]---------------------------
$scriptName   = "deploy"
$logFileName = "$($scriptName).log"

# ---------------------------[ Configuration ]---------------------------
$ErrorActionPreference = 'Stop'

$enableGroupCreation   = $true
$groupNamingAppSuffix  = $true
$groupCreationBlacklist = @(
    "9WZDNCRFJBH4", # Microsoft Photos
    "9PCFS5B6T72H", # Paint
    "9MZ95KL8MR0L", # Snipping Tool
    "9WZDNCRFHVN5", # Windows Calculator
    "9MSMLRH6LZF3", # Windows Notepad
    "9N0DX20HK701", # Windows Terminal
    "9NRX63209R7B", # Outlook (new)
    "9WZDNCRFJ3PZ", # Company Portal
    "XPDP273C0XHQH2" # Adobe Acrobat Reader DC
)

$graphBaseUrl          = "https://graph.microsoft.com/beta"
$graphScopes           = @('DeviceManagementApps.ReadWrite.All', 'Group.ReadWrite.All')

# ---------------------------[ Logging Setup ]---------------------------
$log           = $true
$logDebug      = $false
$logGet        = $true
$logRun        = $true
$enableLogFile = $true

$logFileDirectory = Join-Path $PSScriptRoot 'logs'
$logFile          = Join-Path $logFileDirectory $logFileName

if ($enableLogFile -and -not (Test-Path -Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

# ---------------------------[ Logging Function ]---------------------------
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$message,
        [string]$tag = "Info"
    )

    if (-not $log) { return }

    if (($tag -eq "Debug") -and (-not $logDebug)) { return }
    if (($tag -eq "Get")   -and (-not $logGet))   { return }
    if (($tag -eq "Run")   -and (-not $logRun))   { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start","Get","Run","Info","Success","Error","Debug","End")
    $rawTag    = $tag.Trim()

    if ($tagList -contains $rawTag) {
        $rawTag = $rawTag.PadRight(7)
    }
    else {
        $rawTag = "Error  "
    }

    $color = switch ($rawTag.Trim()) {
        "Start"   { "Cyan" }
        "Get"     { "Blue" }
        "Run"     { "Magenta" }
        "Info"    { "Yellow" }
        "Success" { "Green" }
        "Error"   { "Red" }
        "Debug"   { "DarkYellow" }
        "End"     { "Cyan" }
        default   { "White" }
    }

    $logMessage = "$timestamp [  $rawTag ] $message"

    if ($enableLogFile) {
        try {
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
        }
        catch {
            # Logging must never block script execution
        }
    }

    Write-Host "$timestamp " -NoNewline
    Write-Host "[  " -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host " ] " -NoNewline -ForegroundColor White
    Write-Host "$message"
}

# ---------------------------[ Exit Function ]---------------------------
function Complete-Script {
    param([int]$exitCode)

    $scriptEndTime = Get-Date
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Script execution time: $($duration.ToString('hh\:mm\:ss\.ff'))" -tag "Info"
    Write-Log "Exit Code: $exitCode" -tag "Info"
    Write-Log "======== Script Completed ========" -tag "End"

    exit $exitCode
}

# ---------------------------[ Files and Folders ]---------------------------
$rootDir      = Split-Path -Parent $PSCommandPath
$csvPath      = Join-Path $rootDir 'apps.csv'
$metadataDir  = Join-Path $rootDir 'metadata'
$iconsRoot    = Join-Path $rootDir 'icons'

# ---------------------------[ Auth ]---------------------------
function Test-GraphModulesInstalled {
    $authModule   = Get-Module -ListAvailable 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue | Select-Object -First 1
    $graphModule  = Get-Module -ListAvailable 'Microsoft.Graph' -ErrorAction SilentlyContinue | Select-Object -First 1
    $graphBeta    = Get-Module -ListAvailable 'Microsoft.Graph.Beta' -ErrorAction SilentlyContinue | Select-Object -First 1
    $missing = @()
    if (-not $authModule -and -not $graphModule) { $missing += 'Microsoft.Graph' }
    if (-not $graphBeta) { $missing += 'Microsoft.Graph.Beta' }
    if ($missing.Count -gt 0) {
        throw "Missing Graph module(s): $($missing -join ', '). Install with: Install-Module -Name Microsoft.Graph,Microsoft.Graph.Beta -Scope CurrentUser"
    }
    return $true
}

function Test-GraphScopes {
    param([object]$graphContext)
    if (-not $graphContext) { return $false }
    $currentScopes = @()
    if ($graphContext.Scopes) {
        if ($graphContext.Scopes -is [array]) {
            $currentScopes = $graphContext.Scopes
        } else {
            $currentScopes = @($graphContext.Scopes -split '\s+')
        }
    }
    foreach ($required in $graphScopes) {
        if ($currentScopes -notcontains $required) {
            Write-Log "Scope missing: $required (re-auth required)" -tag "Debug"
            return $false
        }
    }
    return $true
}

function Initialize-GraphConnection {
    Test-GraphModulesInstalled | Out-Null
    if (-not (Get-Module 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue)) {
        Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop
        Write-Log "Loaded Microsoft.Graph.Authentication" -tag "Debug"
    }
    $graphContext = $null
    try {
        $graphContext = Get-MgContext -ErrorAction Stop
    } catch {
        Write-Log "No existing Graph context." -tag "Debug"
    }
    if ($graphContext -and $graphContext.Account -and (Test-GraphScopes -graphContext $graphContext)) {
        $tenantInfo = if ($graphContext.TenantId) { " | TenantId: $($graphContext.TenantId)" } else { '' }
        Write-Log "Graph already connected: $($graphContext.Account)$tenantInfo" -tag "Success"
        return
    }
    if ($graphContext -and -not (Test-GraphScopes -graphContext $graphContext)) {
        Write-Log "Re-authenticating to grant required scopes..." -tag "Run"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Log "Connecting to Graph..." -tag "Run"
    Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    $tenantInfo = if ($graphContext -and $graphContext.TenantId) { " | TenantId: $($graphContext.TenantId)" } else { '' }
    Write-Log "Connected: $($graphContext.Account)$tenantInfo" -tag "Success"
}

# ---------------------------[ Graph Request ]---------------------------
function Invoke-GraphApi {
    param([string]$method, [string]$resource, [object]$body = $null)
    $requestUri = if ($resource -match '^https?://') { $resource } else { "$graphBaseUrl/$($resource.TrimStart('/'))" }
    Write-Log "Graph $method $requestUri" -tag "Debug"
    $invokeParams = @{ Uri = $requestUri; Method = $method }
    if ($body -ne $null) {
        $invokeParams.Body = if ($body -is [string]) { $body } else { $body | ConvertTo-Json -Depth 15 -Compress:$false }
        $invokeParams.ContentType = 'application/json'
    }
    return Invoke-MgGraphRequest @invokeParams
}

# ---------------------------[ Metadata from metadata/<StoreId>.json ]---------------------------
function Get-AppMetadata {
    param([Parameter(Mandatory)][string]$storeId)
    $metadataFile = Join-Path $metadataDir "$storeId.json"
    if (-not (Test-Path -LiteralPath $metadataFile)) { return $null }
    try {
        return Get-Content -LiteralPath $metadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "Failed to parse metadata for $storeId : $($_.Exception.Message)" -tag "Error"
        return $null
    }
}

# ---------------------------[ Icon base64 ]---------------------------
# Resolves icon by: 1) exact match (AppName.png), 2) prefix match (app starts with icon base name, longest wins)
function Get-IconBase64 {
    param([string]$appName, [string]$iconsFolder)
    if (-not (Test-Path -LiteralPath $iconsFolder)) { return $null }
    $allIcons = Get-ChildItem -LiteralPath $iconsFolder -Filter '*.png' -File -ErrorAction SilentlyContinue
    if (-not $allIcons) { return $null }
    $exactMatch = $allIcons | Where-Object { $_.BaseName -eq $appName } | Select-Object -First 1
    if ($exactMatch) {
        Write-Log "Icon: $($exactMatch.Name)" -tag "Get"
        return [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($exactMatch.FullName))
    }
    $prefixMatches = $allIcons | Where-Object { $appName.StartsWith($_.BaseName, [StringComparison]::OrdinalIgnoreCase) }
    if (-not $prefixMatches) { return $null }
    $iconFile = $prefixMatches | Sort-Object { $_.BaseName.Length } -Descending | Select-Object -First 1
    Write-Log "Icon: $($iconFile.Name)" -tag "Get"
    return [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($iconFile.FullName))
}

# ---------------------------[ WinGet Store app body ]---------------------------
function New-WinGetAppBody {
    param(
        [Parameter(Mandatory)][string]$displayName,
        [Parameter(Mandatory)][string]$packageIdentifier,
        [string]$description = '',
        [string]$publisher = '',
        [string]$informationUrl = '',
        [string]$privacyUrl = '',
        [string]$installContext = 'system',
        [string]$owner = '',
        [string]$notes = '',
        [string]$iconBase64 = $null
    )

    $runAsAccount = if ($installContext -eq 'user') { 'User' } else { 'System' }

    if ($informationUrl -and $informationUrl -notmatch '^https?://') { $informationUrl = "https://$informationUrl" }
    if ($privacyUrl -and $privacyUrl -notmatch '^https?://') { $privacyUrl = "https://$privacyUrl" }

    $appBody = @{
        '@odata.type'         = '#microsoft.graph.winGetApp'
        displayName           = $displayName
        description           = $description
        publisher             = $publisher
        developer             = if ($Publisher -eq 'Microsoft Corporation') { 'Microsoft' } else { $Publisher }
        owner                 = $owner
        notes                 = $notes
        packageIdentifier     = $packageIdentifier
        repositoryType        = 'microsoftStore'
        installExperience     = @{ runAsAccount = $runAsAccount }
        isFeatured            = $false
        informationUrl        = $informationUrl
        privacyInformationUrl = $privacyUrl
        roleScopeTagIds       = @()
    }

    if (-not [string]::IsNullOrWhiteSpace($iconBase64)) {
        $appBody.largeIcon = @{
            '@odata.type' = '#microsoft.graph.mimeContent'
            type          = 'image/png'
            value         = $iconBase64
        }
    }

    return $appBody
}

# ---------------------------[ Group creation ]---------------------------
function Get-ResolvedGroupNames {
    param([string]$appName)
    if ($groupNamingAppSuffix) {
        return @{ RQ = "Win - SW - RQ - $appName"; AV = "Win - SW - AV - $appName" }
    }
    return @{ RQ = "Win - SW - $appName - RQ"; AV = "Win - SW - $appName - AV" }
}

function Test-AppGroupBlacklisted {
    param([string]$storeId)
    foreach ($entry in $groupCreationBlacklist) {
        if ($storeId -eq $entry) { return $true }
    }
    return $false
}

function Get-SafeMailNickname {
    param([string]$displayName)
    $safe = ($displayName -replace '[^a-zA-Z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Group' }
    return $safe.Substring(0, [Math]::Min(56, $safe.Length))
}

function Initialize-AppGroup {
    param([string]$displayName)
    $escaped = $displayName -replace "'", "''"
    $filter = [uri]::EscapeDataString("displayName eq '$escaped'")
    $existing = Invoke-GraphApi -method Get -resource "/groups?`$filter=$filter&`$top=1&`$select=id,displayName"
    if ($existing.value -and $existing.value.Count -gt 0) {
        Write-Log "Group exists: $displayName" -tag "Get"
        return $existing.value[0].id
    }
    $mailNick = (Get-SafeMailNickname -displayName $displayName) + [guid]::NewGuid().ToString('N').Substring(0, 8)
    $groupBody = @{
        displayName     = $displayName
        description     = "Intune app assignment group: $displayName"
        mailEnabled     = $false
        mailNickname    = $mailNick
        securityEnabled = $true
        groupTypes      = @()
    }
    Write-Log "Creating group: $displayName" -tag "Run"
    $created = Invoke-GraphApi -method Post -resource '/groups' -body $groupBody
    Write-Log "Created group: $displayName (id: $($created.id))" -tag "Debug"
    return $created.id
}

function Set-AppGroupAssignments {
    param([string]$appId, [string]$displayName)
    $groupNames = Get-ResolvedGroupNames -appName $displayName
    $rqGroupId = Initialize-AppGroup -displayName $groupNames.RQ
    $avGroupId = Initialize-AppGroup -displayName $groupNames.AV
    $assignments = @(
        @{
            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
            target        = @{
                '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                groupId       = $rqGroupId
            }
            intent        = 'required'
            settings      = @{
                '@odata.type'       = '#microsoft.graph.winGetAppAssignmentSettings'
                notifications       = 'hideAll'
                installTimeSettings = $null
                restartSettings     = $null
            }
        }
        @{
            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
            target        = @{
                '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                groupId       = $avGroupId
            }
            intent        = 'available'
            settings      = @{
                '@odata.type'       = '#microsoft.graph.winGetAppAssignmentSettings'
                notifications       = 'showAll'
                installTimeSettings = $null
                restartSettings     = $null
            }
        }
    )
    $assignBody = @{ mobileAppAssignments = $assignments }
    Invoke-GraphApi -method Post -resource "/deviceAppManagement/mobileApps/$appId/assign" -body $assignBody
    Write-Log "Assigned app to groups: RQ=$($groupNames.RQ), AV=$($groupNames.AV)" -tag "Debug"
    Write-Log "Assigned groups to app" -tag "Success"
}

# ---------------------------[ Check if app exists ]---------------------------
function Test-AppExists {
    param([string]$displayName)
    try {
        $escapedName = $displayName -replace "'", "''"
        $odataFilter = [uri]::EscapeDataString("displayName eq '$escapedName'")
        $response = Invoke-GraphApi -method Get -resource "/deviceAppManagement/mobileApps?`$filter=$odataFilter&`$top=1&`$select=id,displayName"
        return ($response.value -and $response.value.Count -gt 0)
    } catch { return $false }
}

# ---------------------------[ Deploy Store App ]---------------------------
function Deploy-StoreApp {
    param(
        [Parameter(Mandatory)][string]$appName,
        [Parameter(Mandatory)][string]$storeId,
        [string]$installContext = 'system'
    )

    if (Test-AppExists -displayName $appName) {
        Write-Log "Skipped (already in Intune): $appName" -tag "Info"
        return 'Skipped'
    }

    Write-Log "Processing: $appName ($storeId) [context=$installContext]" -tag "Info"

    $description    = ''
    $publisher      = ''
    $informationUrl = ''
    $privacyUrl     = ''

    $metadata = Get-AppMetadata -storeId $storeId
    if ($metadata) {
        $description    = if ($metadata.Description) { $metadata.Description } else { '' }
        $publisher      = if ($metadata.Publisher) { $metadata.Publisher } else { '' }
        $informationUrl = if ($metadata.InformationUrl) { $metadata.InformationUrl } elseif ($metadata.PublisherUrl) { $metadata.PublisherUrl } else { '' }
        $privacyUrl     = if ($metadata.PrivacyUrl) { $metadata.PrivacyUrl } else { '' }
        Write-Log "Metadata loaded for: $appName (Publisher='$publisher')" -tag "Debug"
    } else {
        Write-Log "No metadata found for: $appName ($storeId) - run fetch.ps1 on Windows first" -tag "Info"
    }

    $iconBase64 = Get-IconBase64 -appName $appName -iconsFolder $iconsRoot
    if (-not $iconBase64) {
        Write-Log "No icon file retrieved for: $appName" -tag "Info"
    }

    try {
        $appBody = New-WinGetAppBody -displayName $appName -packageIdentifier $storeId `
            -description $description -publisher $publisher `
            -informationUrl $informationUrl -privacyUrl $privacyUrl `
            -installContext $installContext -iconBase64 $iconBase64

        if ($logDebug) {
            $dumpPath = Join-Path $logFileDirectory 'deploy-request-body.json'
            $appBody | ConvertTo-Json -Depth 15 -Compress:$false | Set-Content -Path $dumpPath -Encoding UTF8
            Write-Log "Request body saved: $dumpPath" -tag "Debug"
        }

        Write-Log "Creating Store app: $appName" -tag "Run"
        $createdApp = Invoke-GraphApi -method Post -resource '/deviceAppManagement/mobileApps' -body $appBody
        $createdAppId = $createdApp.id
        Write-Log "Created app id: $createdAppId" -tag "Success"

        if ($enableGroupCreation -and -not (Test-AppGroupBlacklisted -storeId $storeId)) {
            $assignMaxRetries = 3
            $assignDelaySec   = 10
            $assignDone       = $false
            for ($attempt = 1; $attempt -le $assignMaxRetries; $attempt++) {
                try {
                    Set-AppGroupAssignments -appId $createdAppId -displayName $appName
                    $assignDone = $true
                    break
                } catch {
                    $errMsg = $_.Exception.Message
                    $isPublishState = $errMsg -match "PublishingState is not 'Published'"
                    if ($isPublishState -and $attempt -lt $assignMaxRetries) {
                        Write-Log "App not yet published, retrying in ${assignDelaySec}s (attempt $attempt/$assignMaxRetries)..." -tag "Info"
                        Start-Sleep -Seconds $assignDelaySec
                    } else {
                        Write-Log "Group creation/assignment failed (non-fatal): $errMsg" -tag "Error"
                        break
                    }
                }
            }
        } elseif ($enableGroupCreation -and (Test-AppGroupBlacklisted -storeId $storeId)) {
            Write-Log "Skipping group creation (app on blacklist): $appName ($storeId)" -tag "Info"
        }

        return $true
    } catch {
        Write-Log "Deploy failed for $appName : $_" -tag "Error"
        return $false
    }
}

# ---------------------------[ Script Start ]---------------------------
Write-Log "======== Script Started ========" -tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -tag "Info"
Write-Log "Config: graphBaseUrl=$graphBaseUrl | enableGroupCreation=$enableGroupCreation | groupNamingAppSuffix=$groupNamingAppSuffix" -tag "Debug"

# Check metadata directory (produced by fetch.ps1 on Windows)
if (Test-Path -LiteralPath $metadataDir) {
    $metadataFileCount = @(Get-ChildItem -Path $metadataDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Write-Log "Metadata directory found ($metadataFileCount file(s))" -tag "Get"
} else {
    Write-Log "Metadata directory not found - deploying without metadata (run fetch.ps1 on Windows to populate)" -tag "Info"
}

if (-not (Test-Path -LiteralPath $csvPath)) {
    Write-Log "CSV file not found: $csvPath" -tag "Error"
    Complete-Script -exitCode 1
}

try { Initialize-GraphConnection } catch {
    Write-Log "Graph connection failed: $_" -tag "Error"
    Complete-Script -exitCode 1
}

Write-Log "Loading CSV from: $csvPath" -tag "Get"
try {
    $rows = Import-Csv -LiteralPath $csvPath -Delimiter ','
    Write-Log "CSV loaded: $($rows.Count) row(s)" -tag "Debug"
} catch {
    Write-Log "Failed to read CSV: $($_.Exception.Message)" -tag "Error"
    Complete-Script -exitCode 1
}

if (-not $rows -or $rows.Count -eq 0) {
    Write-Log 'CSV contains no rows.' -tag 'Error'
    Complete-Script -exitCode 1
}

$totalApps     = $rows.Count
$deployedCount = 0
$failedCount   = 0
$skippedCount  = 0

Write-Log "Processing $totalApps app(s) from CSV" -tag "Info"

foreach ($row in $rows) {
    $appName        = ($row.ApplicationName).ToString().Trim()
    $storeId        = ($row.StoreId).ToString().Trim()
    $installContext = ($row.InstallContext).ToString().Trim().ToLower()

    if ([string]::IsNullOrWhiteSpace($appName) -or [string]::IsNullOrWhiteSpace($storeId)) {
        Write-Log 'Skipping row with missing ApplicationName or StoreId.' -tag 'Error'
        $failedCount++
        continue
    }

    if ($installContext -notin @('system', 'user')) {
        Write-Log "Invalid InstallContext '$installContext' for $appName - defaulting to 'system'" -tag "Info"
        $installContext = 'system'
    }

    $result = Deploy-StoreApp -appName $appName -storeId $storeId -installContext $installContext
    if ($result -eq $true)        { $deployedCount++ }
    elseif ($result -eq 'Skipped') { $skippedCount++ }
    else                          { $failedCount++ }
}

Write-Log "Deploy summary: $deployedCount succeeded, $skippedCount skipped, $failedCount failed (total: $totalApps)" -tag "Info"
Complete-Script -exitCode $(if ($failedCount -gt 0) { 1 } else { 0 })
