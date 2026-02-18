# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

# ---------------------------[ Script Name ]---------------------------
$scriptName   = "fetch"
$logFileName = "$($scriptName).log"

# ---------------------------[ Configuration ]---------------------------
$ErrorActionPreference = 'Stop'

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

# ---------------------------[ Winget Localization ]---------------------------
function ConvertFrom-WingetLocalizedOutput {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$rawOutput)

    $langPath = Join-Path (Join-Path $PSScriptRoot 'jsons') 'language.json'
    if (-not (Test-Path -LiteralPath $langPath)) {
        Write-Log "language.json not found, using raw output" -tag "Debug"
        return $rawOutput
    }

    try {
        $lang = Get-Content -LiteralPath $langPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "Failed to load language.json: $($_.Exception.Message)" -tag "Debug"
        return $rawOutput
    }

    $culture = [System.Globalization.CultureInfo]::CurrentUICulture.Name
    $localeKey = $null
    if ($lang.locales.PSObject.Properties.Name -contains $culture) {
        $localeKey = $culture
    } elseif ($culture -match '^([a-z]{2})-') {
        $baseCulture = $Matches[1]
        $localeKey = $lang.locales.PSObject.Properties.Name | Where-Object { $_ -like "$baseCulture-*" } | Select-Object -First 1
    }

    if (-not $localeKey) {
        Write-Log "No locale mapping for '$culture', using raw output" -tag "Debug"
        return $rawOutput
    }

    $locale = $lang.locales.$localeKey
    $result = $rawOutput

    $noPkgPatterns = @($lang.english.noPackageSubstrings)
    if ($locale.noPackageSubstrings -and $locale.noPackageSubstrings.Count -gt 0) {
        $noPkgPatterns = @($locale.noPackageSubstrings)
    }
    foreach ($pat in $noPkgPatterns) {
        if ($result -like "*$pat*") {
            return $null
        }
    }

    $foundWord = $locale.foundWord
    if ($foundWord -and $foundWord -ne 'Found') {
        $lines = $result -split "`r?`n"
        if ($lines.Count -gt 0 -and $lines[0] -match "^$([regex]::Escape($foundWord))\s+") {
            $lines[0] = $lines[0] -replace "^$([regex]::Escape($foundWord))\s+", 'Found '
            $result = $lines -join "`n"
        }
    }

    if ($locale.labels -and $locale.labels.PSObject.Properties) {
        $locale.labels.PSObject.Properties | Sort-Object { $_.Name.Length } -Descending | ForEach-Object {
            $localized = $_.Name
            $english = $_.Value
            if ($localized -and $english -and $localized -ne $english) {
                $result = $result -replace ([regex]::Escape($localized) + ':'), ($english + ':')
            }
        }
    }

    return $result
}

# ---------------------------[ Winget Show ]---------------------------
function Invoke-WingetShowRaw {
    param([string]$wingetId)
    $wingetArgs = @('show', '--id', $wingetId)
    $prevEnc = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    try {
        $out = & winget @wingetArgs 2>&1 | Out-String
        return @{ Output = $out; ExitCode = $LASTEXITCODE }
    } finally {
        [Console]::OutputEncoding = $prevEnc
    }
}

# ---------------------------[ Parse Winget Show Output ]---------------------------
function ConvertFrom-WingetShowOutput {
    param([string]$normalizedOutput)
    $result = @{ Obj = [ordered]@{ }; HasInstaller = $false }
    if ([string]::IsNullOrWhiteSpace($normalizedOutput)) { return $result }

    $lines = $normalizedOutput -split "`r?`n"
    $obj   = [ordered]@{ }
    $currentKey = $null
    $currentValue = [System.Collections.ArrayList]::new()
    $currentSection = $null
    $installersList = [System.Collections.ArrayList]::new()
    $normalizeKey = { param([string]$k) ($k -replace '\s+', '').Trim() }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($i -eq 0 -and $line -match 'Found\s+(.+?)\s+\[(.+?)\]') {
            $obj['Name'] = $Matches[1].Trim()
            $obj['Id']   = $Matches[2].Trim()
            continue
        }
        # Continuation lines for multi-line fields (e.g. Description)
        if ($currentSection -eq 'MultiLine' -and $line -match '^\s{2,}(.+)$') {
            [void]$currentValue.Add($Matches[1].Trim())
            continue
        }
        if ($line -match '^([A-Za-z][A-Za-z0-9\s\-]*):\s*(.*)$' -and $line -notmatch '^\s{2,}') {
            if ($currentKey -and $currentValue.Count -gt 0) {
                $val = if ($currentValue.Count -eq 1) { $currentValue[0] } else { ($currentValue -join "`n") }
                $obj[$currentKey] = $val
            }
            $key = $Matches[1].Trim() -replace '\s+', ' '
            $val = $Matches[2].Trim()
            $currentKey = $null
            $currentValue = [System.Collections.ArrayList]::new()
            if ($key -eq 'Release Notes') {
                $currentKey = 'ReleaseNotes'
                $currentSection = 'ReleaseNotes'
                if ($val) { [void]$currentValue.Add($val) }
            } elseif ($key -eq 'Tags') {
                $currentKey = 'Tags'
                $currentSection = 'Tags'
                if ($val) { [void]$currentValue.Add($val) }
            } elseif ($key -eq 'Documentation') {
                $currentSection = 'Documentation'
                if (-not $obj['Documentation']) { $obj['Documentation'] = [ordered]@{ } }
            } elseif ($key -match '^Installer(\s+\d+)?$') {
                $currentSection = 'Installer'
                $singleInstaller = [ordered]@{ }
                if ($obj['Installer'] -is [System.Collections.Specialized.OrderedDictionary]) {
                    [void]$installersList.Add($obj['Installer'])
                    $obj.Remove('Installer')
                }
                [void]$installersList.Add($singleInstaller)
                $obj['Installer'] = $singleInstaller
            } else {
                $normKey = & $normalizeKey $key
                if ($normKey -and ($val -or $val -eq '')) {
                    $currentKey = $normKey
                    $currentSection = 'MultiLine'
                    [void]$currentValue.Add($val)
                }
            }
            continue
        }
        if ($line -match '^\s{2,}([^:]+):\s*(.*)$') {
            $subKey = & $normalizeKey $Matches[1].Trim()
            $subVal = $Matches[2].Trim()
            if ($currentSection -eq 'Documentation' -and $obj['Documentation'] -is [System.Collections.Specialized.OrderedDictionary]) {
                $obj['Documentation'][$subKey] = $subVal
            } elseif ($currentSection -eq 'Installer' -and $obj['Installer'] -is [System.Collections.Specialized.OrderedDictionary]) {
                $obj['Installer'][$subKey] = $subVal
            } elseif ($currentSection -eq 'ReleaseNotes') {
                [void]$currentValue.Add($line.Trim())
            } elseif ($currentSection -eq 'Tags') {
                [void]$currentValue.Add($subKey)
            }
            continue
        }
        if ($line -match '^\s{2,}(.+)$' -and $currentSection -in 'ReleaseNotes','Tags') {
            [void]$currentValue.Add($Matches[1].Trim())
        }
    }

    if ($currentKey -and $currentValue.Count -gt 0) {
        $val = if ($currentValue.Count -eq 1) { $currentValue[0] } else { ($currentValue -join "`n") }
        $obj[$currentKey] = $val
    }
    if ($installersList.Count -gt 1) {
        $obj['Installers'] = @($installersList.ToArray())
        $obj['Installer'] = $installersList[0]
    }
    $hasInstaller = ($obj['Installer'] -and $obj['Installer'].PSObject.Properties.Count -gt 0) -or
                    ($obj['Installers'] -and $obj['Installers'].Count -gt 0)
    return @{ Obj = $obj; HasInstaller = $hasInstaller }
}

# ---------------------------[ Script Start ]---------------------------
Write-Log "======== Script Started ========" -tag "Start"
Write-Log "ComputerName: $env:COMPUTERNAME | User: $env:USERNAME | Script: $scriptName" -tag "Info"

if (-not (Test-Path -LiteralPath $csvPath)) {
    Write-Log "CSV file not found: $csvPath" -tag "Error"
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

# Ensure metadata directory exists
if (-not (Test-Path -Path $metadataDir)) {
    New-Item -ItemType Directory -Path $metadataDir -Force | Out-Null
    Write-Log "Created metadata directory: $metadataDir" -tag "Info"
}

$existingCount = @(Get-ChildItem -Path $metadataDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
Write-Log "Metadata folder contains $existingCount existing file(s)" -tag "Get"

$totalApps   = $rows.Count
$fetchedCount = 0
$skippedCount = 0
$failedCount  = 0

Write-Log "Fetching metadata for $totalApps app(s)" -tag "Info"

foreach ($row in $rows) {
    $appName = ($row.ApplicationName).ToString().Trim()
    $storeId = ($row.StoreId).ToString().Trim()

    if ([string]::IsNullOrWhiteSpace($appName) -or [string]::IsNullOrWhiteSpace($storeId)) {
        Write-Log 'Skipping row with missing ApplicationName or StoreId.' -tag 'Error'
        $failedCount++
        continue
    }

    $metadataFile = Join-Path $metadataDir "$storeId.json"
    if (Test-Path -LiteralPath $metadataFile) {
        Write-Log "Skipped (already fetched): $appName ($storeId)" -tag "Info"
        $skippedCount++
        continue
    }

    Write-Log "Fetching: $appName ($storeId)" -tag "Get"

    $run = Invoke-WingetShowRaw -wingetId $storeId
    if ($run.ExitCode -ne 0 -or [string]::IsNullOrWhiteSpace($run.Output)) {
        Write-Log "winget show returned no output for '$storeId' (exit: $($run.ExitCode))" -tag "Error"
        $failedCount++
        continue
    }

    $normalized = ConvertFrom-WingetLocalizedOutput -rawOutput $run.Output
    if ($null -eq $normalized -or $normalized -match 'No package found|No applicable package|No applicable installer') {
        Write-Log "No package found for '$storeId'" -tag "Error"
        $failedCount++
        continue
    }

    $parsed = ConvertFrom-WingetShowOutput -normalizedOutput $normalized
    $obj = $parsed.Obj

    $descriptionRaw = $obj['Description']
    if ($descriptionRaw -is [array]) { $descriptionRaw = $descriptionRaw -join "`n" }
    $descriptionStr = if ($descriptionRaw) { $descriptionRaw.ToString().Trim() } else { '' }

    $appMetadata = [ordered]@{
        Name           = $appName
        Description    = $descriptionStr
        Publisher      = if ($obj['Publisher']) { $obj['Publisher'] } else { '' }
        PublisherUrl   = if ($obj['PublisherUrl']) { $obj['PublisherUrl'] } else { $null }
        InformationUrl = if ($obj['PackageUrl']) { $obj['PackageUrl'] } elseif ($obj['Homepage']) { $obj['Homepage'] } else { $null }
        PrivacyUrl     = if ($obj['PrivacyUrl']) { $obj['PrivacyUrl'] } else { $null }
        FetchedAt      = (Get-Date -Format 'o')
    }

    try {
        $appMetadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $metadataFile -Encoding UTF8
        Write-Log "Saved: $metadataFile" -tag "Success"
    } catch {
        Write-Log "Failed to write $($metadataFile): $($_.Exception.Message)" -tag "Error"
        $failedCount++
        continue
    }

    Write-Log "Fetched: $appName | Publisher='$($appMetadata.Publisher)'" -tag "Success"
    $fetchedCount++
}

$totalMetadata = @(Get-ChildItem -Path $metadataDir -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
Write-Log "Metadata folder contains $totalMetadata total file(s)" -tag "Success"

Write-Log "Fetch summary: $fetchedCount fetched, $skippedCount skipped, $failedCount failed (total: $totalApps)" -tag "Info"
Complete-Script -exitCode $(if ($failedCount -gt 0) { 1 } else { 0 })
