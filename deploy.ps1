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
        [string]$Message,
        [string]$Tag = "Info"
    )

    if (-not $log) { return }

    if (($Tag -eq "Debug") -and (-not $logDebug)) { return }
    if (($Tag -eq "Get")   -and (-not $logGet))   { return }
    if (($Tag -eq "Run")   -and (-not $logRun))   { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tagList   = @("Start","Get","Run","Info","Success","Error","Debug","End")
    $rawTag    = $Tag.Trim()

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

    $logMessage = "$timestamp [  $rawTag ] $Message"

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
    Write-Host "$Message"
}

# ---------------------------[ Exit Function ]---------------------------
function Complete-Script {
    param([int]$ExitCode)

    $scriptEndTime = Get-Date
    $duration      = $scriptEndTime - $scriptStartTime

    Write-Log "Script execution time: $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag "Info"
    Write-Log "Exit Code: $ExitCode" -Tag "Info"
    Write-Log "======== Script Completed ========" -Tag "End"

    exit $ExitCode
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
    param([object]$GraphContext)
    if (-not $GraphContext) { return $false }
    $currentScopes = @()
    if ($GraphContext.Scopes) {
        if ($GraphContext.Scopes -is [array]) {
            $currentScopes = $GraphContext.Scopes
        } else {
            $currentScopes = @($GraphContext.Scopes -split '\s+')
        }
    }
    foreach ($required in $graphScopes) {
        if ($currentScopes -notcontains $required) {
            Write-Log "Scope missing: $required (re-auth required)" -Tag "Debug"
            return $false
        }
    }
    return $true
}

function Initialize-GraphConnection {
    Test-GraphModulesInstalled | Out-Null
    if (-not (Get-Module 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue)) {
        Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop
        Write-Log "Loaded Microsoft.Graph.Authentication" -Tag "Debug"
    }
    $graphContext = $null
    try {
        $graphContext = Get-MgContext -ErrorAction Stop
    } catch {
        Write-Log "No existing Graph context." -Tag "Debug"
    }
    if ($graphContext -and $graphContext.Account -and (Test-GraphScopes -GraphContext $graphContext)) {
        $tenantInfo = if ($graphContext.TenantId) { " | TenantId: $($graphContext.TenantId)" } else { '' }
        Write-Log "Graph already connected: $($graphContext.Account)$tenantInfo" -Tag "Success"
        return
    }
    if ($graphContext -and -not (Test-GraphScopes -GraphContext $graphContext)) {
        Write-Log "Re-authenticating to grant required scopes..." -Tag "Run"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    Write-Log "Connecting to Graph..." -Tag "Run"
    Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
    $tenantInfo = if ($graphContext -and $graphContext.TenantId) { " | TenantId: $($graphContext.TenantId)" } else { '' }
    Write-Log "Connected: $($graphContext.Account)$tenantInfo" -Tag "Success"
}

# ---------------------------[ Graph Request ]---------------------------
function Invoke-GraphApi {
    param([string]$Method, [string]$Resource, [object]$Body = $null)
    $requestUri = if ($Resource -match '^https?://') { $Resource } else { "$graphBaseUrl/$($Resource.TrimStart('/'))" }
    Write-Log "Graph $Method $requestUri" -Tag "Debug"
    $invokeParams = @{ Uri = $requestUri; Method = $Method }
    if ($Body -ne $null) {
        $invokeParams.Body = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 15 -Compress:$false }
        $invokeParams.ContentType = 'application/json'
    }
    return Invoke-MgGraphRequest @invokeParams
}

# ---------------------------[ Metadata from metadata/<StoreId>.json ]---------------------------
function Get-AppMetadata {
    param([Parameter(Mandatory)][string]$StoreId)
    $metadataFile = Join-Path $metadataDir "$StoreId.json"
    if (-not (Test-Path -LiteralPath $metadataFile)) { return $null }
    try {
        return Get-Content -LiteralPath $metadataFile -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "Failed to parse metadata for $StoreId : $($_.Exception.Message)" -Tag "Error"
        return $null
    }
}

# ---------------------------[ Icon base64 ]---------------------------
# Resolves icon by: 1) exact match (AppName.png), 2) prefix match (app starts with icon base name, longest wins)
function Get-IconBase64 {
    param([string]$AppName, [string]$IconsFolder)
    if (-not (Test-Path -LiteralPath $IconsFolder)) { return $null }
    $allIcons = Get-ChildItem -LiteralPath $IconsFolder -Filter '*.png' -File -ErrorAction SilentlyContinue
    if (-not $allIcons) { return $null }
    $exactMatch = $allIcons | Where-Object { $_.BaseName -eq $AppName } | Select-Object -First 1
    if ($exactMatch) {
        Write-Log "Icon: $($exactMatch.Name)" -Tag "Get"
        return [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($exactMatch.FullName))
    }
    $prefixMatches = $allIcons | Where-Object { $AppName.StartsWith($_.BaseName, [StringComparison]::OrdinalIgnoreCase) }
    if (-not $prefixMatches) { return $null }
    $iconFile = $prefixMatches | Sort-Object { $_.BaseName.Length } -Descending | Select-Object -First 1
    Write-Log "Icon: $($iconFile.Name)" -Tag "Get"
    return [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($iconFile.FullName))
}

# ---------------------------[ WinGet Store app body ]---------------------------
function New-WinGetAppBody {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$PackageIdentifier,
        [string]$Description = '',
        [string]$Publisher = '',
        [string]$InformationUrl = '',
        [string]$PrivacyUrl = '',
        [string]$InstallContext = 'system',
        [string]$Owner = '',
        [string]$Notes = '',
        [string]$IconBase64 = $null
    )

    $runAsAccount = if ($InstallContext -eq 'user') { 'User' } else { 'System' }

    $body = @{
        '@odata.type'         = '#microsoft.graph.winGetApp'
        displayName           = $DisplayName
        description           = $Description
        publisher             = $Publisher
        developer             = if ($Publisher -eq 'Microsoft Corporation') { 'Microsoft' } else { $Publisher }
        owner                 = $Owner
        notes                 = $Notes
        packageIdentifier     = $PackageIdentifier
        repositoryType        = 'microsoftStore'
        installExperience     = @{ runAsAccount = $runAsAccount }
        isFeatured            = $false
        informationUrl        = $InformationUrl
        privacyInformationUrl = $PrivacyUrl
        roleScopeTagIds       = @()
    }

    if (-not [string]::IsNullOrWhiteSpace($IconBase64)) {
        $body.largeIcon = @{
            '@odata.type' = '#microsoft.graph.mimeContent'
            type          = 'image/png'
            value         = $IconBase64
        }
    }

    return $body
}

# ---------------------------[ Group creation ]---------------------------
function Get-ResolvedGroupNames {
    param([string]$AppName)
    if ($groupNamingAppSuffix) {
        return @{ RQ = "Win - SW - RQ - $AppName"; AV = "Win - SW - AV - $AppName" }
    }
    return @{ RQ = "Win - SW - $AppName - RQ"; AV = "Win - SW - $AppName - AV" }
}

function Test-AppGroupBlacklisted {
    param([string]$StoreId)
    foreach ($entry in $groupCreationBlacklist) {
        if ($StoreId -eq $entry) { return $true }
    }
    return $false
}

function Get-SafeMailNickname {
    param([string]$DisplayName)
    $safe = ($DisplayName -replace '[^a-zA-Z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Group' }
    return $safe.Substring(0, [Math]::Min(56, $safe.Length))
}

function Get-OrCreateGroup {
    param([string]$DisplayName)
    $escaped = $DisplayName -replace "'", "''"
    $filter = [uri]::EscapeDataString("displayName eq '$escaped'")
    $existing = Invoke-GraphApi -Method Get -Resource "/groups?`$filter=$filter&`$top=1&`$select=id,displayName"
    if ($existing.value -and $existing.value.Count -gt 0) {
        Write-Log "Group exists: $DisplayName" -Tag "Get"
        return $existing.value[0].id
    }
    $mailNick = (Get-SafeMailNickname -DisplayName $DisplayName) + [guid]::NewGuid().ToString('N').Substring(0, 8)
    $groupBody = @{
        displayName     = $DisplayName
        description     = "Intune app assignment group: $DisplayName"
        mailEnabled     = $false
        mailNickname    = $mailNick
        securityEnabled = $true
        groupTypes      = @()
    }
    Write-Log "Creating group: $DisplayName" -Tag "Run"
    $created = Invoke-GraphApi -Method Post -Resource '/groups' -Body $groupBody
    Write-Log "Created group: $DisplayName (id: $($created.id))" -Tag "Debug"
    return $created.id
}

function Set-AppGroupAssignments {
    param([string]$AppId, [string]$DisplayName)
    $groupNames = Get-ResolvedGroupNames -AppName $DisplayName
    $rqGroupId = Get-OrCreateGroup -DisplayName $groupNames.RQ
    $avGroupId = Get-OrCreateGroup -DisplayName $groupNames.AV
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
    Invoke-GraphApi -Method Post -Resource "/deviceAppManagement/mobileApps/$AppId/assign" -Body $assignBody
    Write-Log "Assigned app to groups: RQ=$($groupNames.RQ), AV=$($groupNames.AV)" -Tag "Debug"
    Write-Log "Assigned groups to app" -Tag "Success"
}

# ---------------------------[ Check if app exists ]---------------------------
function Test-AppExists {
    param([string]$DisplayName)
    try {
        $escapedName = $DisplayName -replace "'", "''"
        $odataFilter = [uri]::EscapeDataString("displayName eq '$escapedName'")
        $response = Invoke-GraphApi -Method Get -Resource "/deviceAppManagement/mobileApps?`$filter=$odataFilter&`$top=1&`$select=id,displayName"
        return ($response.value -and $response.value.Count -gt 0)
    } catch { return $false }
}

# ---------------------------[ Deploy Store App ]---------------------------
function Deploy-StoreApp {
    param(
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$StoreId,
        [string]$InstallContext = 'system'
    )

    if (Test-AppExists -DisplayName $AppName) {
        Write-Log "Skipped (already in Intune): $AppName" -Tag "Info"
        return 'Skipped'
    }

    Write-Log "Processing: $AppName ($StoreId) [context=$InstallContext]" -Tag "Info"

    $description    = ''
    $publisher      = ''
    $informationUrl = ''
    $privacyUrl     = ''

    $metadata = Get-AppMetadata -StoreId $StoreId
    if ($metadata) {
        $description    = if ($metadata.Description) { $metadata.Description } else { '' }
        $publisher      = if ($metadata.Publisher) { $metadata.Publisher } else { '' }
        $informationUrl = if ($metadata.InformationUrl) { $metadata.InformationUrl } elseif ($metadata.PublisherUrl) { $metadata.PublisherUrl } else { '' }
        $privacyUrl     = if ($metadata.PrivacyUrl) { $metadata.PrivacyUrl } else { '' }
        Write-Log "Metadata loaded for: $AppName (Publisher='$publisher')" -Tag "Debug"
    } else {
        Write-Log "No metadata found for: $AppName ($StoreId) - run fetch.ps1 on Windows first" -Tag "Info"
    }

    $iconBase64 = Get-IconBase64 -AppName $AppName -IconsFolder $iconsRoot
    if (-not $iconBase64) {
        Write-Log "No icon file retrieved for: $AppName" -Tag "Info"
    }

    try {
        $appBody = New-WinGetAppBody -DisplayName $AppName -PackageIdentifier $StoreId `
            -Description $description -Publisher $publisher `
            -InformationUrl $informationUrl -PrivacyUrl $privacyUrl `
            -InstallContext $InstallContext -IconBase64 $iconBase64

        if ($logDebug) {
            $dumpPath = Join-Path $logFileDirectory 'deploy-request-body.json'
            $appBody | ConvertTo-Json -Depth 15 -Compress:$false | Set-Content -Path $dumpPath -Encoding UTF8
            Write-Log "Request body saved: $dumpPath" -Tag "Debug"
        }

        Write-Log "Creating Store app: $AppName" -Tag "Run"
        $createdApp = Invoke-GraphApi -Method Post -Resource '/deviceAppManagement/mobileApps' -Body $appBody
        $appId = $createdApp.id
        Write-Log "Created app id: $appId" -Tag "Success"

        if ($enableGroupCreation -and -not (Test-AppGroupBlacklisted -StoreId $StoreId)) {
            try {
                Set-AppGroupAssignments -AppId $appId -DisplayName $AppName
            } catch {
                Write-Log "Group creation/assignment failed (non-fatal): $_" -Tag "Error"
            }
        } elseif ($enableGroupCreation -and (Test-AppGroupBlacklisted -StoreId $StoreId)) {
            Write-Log "Skipping group creation (app on blacklist): $AppName ($StoreId)" -Tag "Info"
        }

        return $true
    } catch {
        Write-Log "Deploy failed for $AppName : $_" -Tag "Error"
        return $false
    }
}

# ---------------------------[ Script Start ]---------------------------
Write-Log "======== Script Started ========" -Tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -Tag "Info"
Write-Log "Config: graphBaseUrl=$graphBaseUrl | enableGroupCreation=$enableGroupCreation | groupNamingAppSuffix=$groupNamingAppSuffix" -Tag "Debug"

# Check metadata directory (produced by fetch.ps1 on Windows)
if (Test-Path -LiteralPath $metadataDir) {
    $metadataFileCount = @(Get-ChildItem -Path $metadataDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
    Write-Log "Metadata directory found ($metadataFileCount file(s))" -Tag "Get"
} else {
    Write-Log "Metadata directory not found - deploying without metadata (run fetch.ps1 on Windows to populate)" -Tag "Info"
}

if (-not (Test-Path -LiteralPath $csvPath)) {
    Write-Log "CSV file not found: $csvPath" -Tag "Error"
    Complete-Script -ExitCode 1
}

try { Initialize-GraphConnection } catch {
    Write-Log "Graph connection failed: $_" -Tag "Error"
    Complete-Script -ExitCode 1
}

Write-Log "Loading CSV from: $csvPath" -Tag "Get"
try {
    $rows = Import-Csv -LiteralPath $csvPath -Delimiter ','
    Write-Log "CSV loaded: $($rows.Count) row(s)" -Tag "Debug"
} catch {
    Write-Log "Failed to read CSV: $($_.Exception.Message)" -Tag "Error"
    Complete-Script -ExitCode 1
}

if (-not $rows -or $rows.Count -eq 0) {
    Write-Log 'CSV contains no rows.' -Tag 'Error'
    Complete-Script -ExitCode 1
}

$totalApps     = $rows.Count
$deployedCount = 0
$failedCount   = 0
$skippedCount  = 0

Write-Log "Processing $totalApps app(s) from CSV" -Tag "Info"

foreach ($row in $rows) {
    $appName        = ($row.ApplicationName).ToString().Trim()
    $storeId        = ($row.StoreId).ToString().Trim()
    $installContext = ($row.InstallContext).ToString().Trim().ToLower()

    if ([string]::IsNullOrWhiteSpace($appName) -or [string]::IsNullOrWhiteSpace($storeId)) {
        Write-Log 'Skipping row with missing ApplicationName or StoreId.' -Tag 'Error'
        $failedCount++
        continue
    }

    if ($installContext -notin @('system', 'user')) {
        Write-Log "Invalid InstallContext '$installContext' for $appName - defaulting to 'system'" -Tag "Info"
        $installContext = 'system'
    }

    $result = Deploy-StoreApp -AppName $appName -StoreId $storeId -InstallContext $installContext
    if ($result -eq $true)        { $deployedCount++ }
    elseif ($result -eq 'Skipped') { $skippedCount++ }
    else                          { $failedCount++ }
}

Write-Log "Deploy summary: $deployedCount succeeded, $skippedCount skipped, $failedCount failed (total: $totalApps)" -Tag "Info"
Complete-Script -ExitCode $(if ($failedCount -gt 0) { 1 } else { 0 })
