<#
.SYNOPSIS
Publishes MPPs to Altus from the current location, under the Files subfolder.

.DESCRIPTION
Given a root folder, this script finds files to publish.
- If -All is supplied, it publishes **all files** (takes precedence over -NamePattern).
- Otherwise, it publishes files matching -NamePattern (default: Project_*_published.mpp).
- Files are ordered by ProjectLastPublishedDate (oldest first) using paired JSON metadata files.
- Use -DryRun to preview the publish order without actually publishing.

.PARAMETER Path
(Optional) The folder to search for MPP files. Defaults to 'Files' subfolder of the script location.

.PARAMETER All
(Optional) Switch. If present, publish **all .mpp files** under -Root. This overrides -NamePattern.

.PARAMETER NamePattern
(Optional) A name pattern (wildcards supported) of files to publish when -All is not provided.
Defaults to 'Project_*_published.mpp'.

.PARAMETER RetryCount
(Optional) Number of times to retry a failed file operation.
Defaults to 3.

.PARAMETER DelayBetweenFiles
(Optional) Delay in seconds between processing files to allow MS Project to stabilize.
Defaults to 2 seconds.

.PARAMETER RetryDelay
(Optional) Delay in seconds before retrying a failed operation.
Defaults to 5 seconds.

.PARAMETER DateFallback
(Optional) Fallback strategy when JSON date is missing or invalid:
  - 'EndOfQueue' (default): Push files without valid dates to the end of the publish queue.
  - 'UseMppWriteTime': Use the MPP file's LastWriteTime as the sort date.

.PARAMETER DryRun
(Optional) Switch. Preview the publish order without actually publishing any files.

.PARAMETER BaselineNumber
(Optional) The baseline number to publish (0-10). When specified with -IsReportable, baseline publishing will be enabled after each project publish.

.PARAMETER IsReportable
(Optional) Whether the baseline should be reportable. When specified with -BaselineNumber, baseline publishing will be enabled.

.EXAMPLE
.\Publish-MPPsToAltus.ps1
Search Files folder recursively for 'Project_*_published.mpp' and publish matches ordered by ProjectLastPublishedDate.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -All
Publish **all .mpp files** under Files folder (overrides -NamePattern).

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -NamePattern '*.mpp'
Only publish *.mpp files under Files folder.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -DryRun
Preview the publish order without actually publishing any files.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -DateFallback UseMppWriteTime
Use MPP file write time when JSON metadata is missing or invalid.

.\Publish-MPPsToAltus.ps1 -All -DelayBetweenFiles 5
Publish **all .mpp files** with 5 second delays between files.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -NamePattern '*.mpp' -RetryCount 5 -RetryDelay 10
Publish *.mpp files with 5 retries and 10 second delays between retries.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -All -ComRestartInterval 25 -RetryCount 5
Publish all files with proactive COM restart every 25 files and 5 retry attempts.

.EXAMPLE
.\Publish-MPPsToAltus.ps1 -BaselineNumber 0 -IsReportable $true
Publish files and also publish baseline 0 as reportable after each project publish.

.NOTES
Run:
  Get-Help .\Publish-MPPsToAltus.ps1 -Full
  Get-Help .\Publish-MPPsToAltus.ps1 -Examples

Features:
- Garbage Collection: Automatic GC every 5 files and before retries to manage COM resources
- Adaptive Delays: Delays automatically decrease after consecutive successes and increase after failures
- Smart Retries: Exceptions trigger retries with the configured RetryCount; logical publish failures are not retried
- COM Restart: Proactive restart every N files (default 30) and reactive restart on RPC errors

For large batches (50+ files), consider:
- Setting ComRestartInterval to 25-40 based on system stability
- Increasing DelayBetweenFiles to 5-10 seconds
- Increasing RetryDelay to 10-15 seconds
#>


[CmdletBinding()]
param(
    [Parameter(HelpMessage = "The folder to search for MPP files. Defaults to 'Files' subfolder of the script location.")]
    [string]$Path,

    [Parameter(HelpMessage = "Publish all .mpp files (overrides -NamePattern).")]
    [switch]$All,

    [Parameter(HelpMessage = "Wildcard pattern of files to publish when -All is not set.")]
    [string]$NamePattern = 'Project_*_published.mpp',

    [Parameter(HelpMessage = "Fallback strategy when JSON date is missing or invalid: 'EndOfQueue' (default) or 'UseMppWriteTime'.")]
    [ValidateSet('EndOfQueue', 'UseMppWriteTime')]
    [string]$DateFallback = 'EndOfQueue',

    [Parameter(HelpMessage = "Preview the publish order without actually publishing.")]
    [switch]$DryRun,

    [Parameter(HelpMessage = "The baseline number to publish (0-10). Required with -IsReportable to enable baseline publishing.")]
    [ValidateRange(0, 10)]
    [int]$BaselineNumber,

    [Parameter(HelpMessage = "Whether the baseline should be reportable. Required with -BaselineNumber to enable baseline publishing.")]
    [bool]$IsReportable,

    [Parameter(HelpMessage = "Number of retry attempts for failed operations.")]
    [int]$RetryCount = 3,

    [Parameter(HelpMessage = "Delay in seconds between processing files.")]
    [int]$DelayBetweenFiles = 2,

    [Parameter(HelpMessage = "Delay in seconds before retrying a failed operation.")]
    [int]$RetryDelay = 5,

    [Parameter(HelpMessage = "Number of files to process before proactively restarting COM connection (0 = disabled).")]
    [int]$ComRestartInterval = 30
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

$ErrorActionPreference = 'Stop';

# Validate baseline parameters - both must be provided together to enable baseline publishing
$baselinePublishEnabled = $PSBoundParameters.ContainsKey('BaselineNumber') -and $PSBoundParameters.ContainsKey('IsReportable')

if ($baselinePublishEnabled) {
    Write-Host "Baseline publishing enabled: Baseline $BaselineNumber, Reportable: $IsReportable" -ForegroundColor Cyan
} elseif ($PSBoundParameters.ContainsKey('BaselineNumber') -or $PSBoundParameters.ContainsKey('IsReportable')) {
    throw "Both -BaselineNumber and -IsReportable must be specified together for baseline publishing."
}

# Validate Common.ps1 exists before dot-sourcing
$commonPath = "$PSScriptRoot\Common.ps1"
if (-not (Test-Path $commonPath)) {
    throw "Required file not found: $commonPath"
}
. $commonPath

if ($PSVersionTable.PSVersion.Major -gt 5) {
    Write-Host "This script does not support PowerShell versions greater than 5.0. Exiting." -ForegroundColor Red
    exit 1
}

$effectivePattern = if ($All) { '*.mpp' } else { $NamePattern }

Write-Verbose "All: $All"
Write-Verbose "Effective pattern: $effectivePattern"
Write-Verbose "DateFallback: $DateFallback"
Write-Verbose "DryRun: $DryRun"
Write-Verbose "Retry Count: $RetryCount"
Write-Verbose "Delay Between Files: $DelayBetweenFiles seconds"
Write-Verbose "Retry Delay: $RetryDelay seconds"
Write-Verbose "COM Restart Interval: $ComRestartInterval files"

# Determine the files path: use -Path if provided, else default to 'Files' subfolder
if ($Path) {
    $filesPath = $Path
    if (-not (Test-Path $filesPath)) {
        Write-Host "Specified path '$filesPath' does not exist. Exiting." -ForegroundColor Red
        exit 1
    }
} else {
    $filesPath = Join-Path $PSScriptRoot "Files"
}

# Resolve to absolute path to ensure consistent path comparison
$filesPath = (Resolve-Path $filesPath).Path

$files = Get-ChildItem -Path $filesPath -File -Recurse -Filter $effectivePattern -ErrorAction Stop

# Exclude files from the Published subfolder (and its subdirectories)
$publishedFolderPath = Join-Path $filesPath "Published"
# Ensure the path ends with a separator to avoid partial matches (e.g. matching "Published_Backup")
if (-not $publishedFolderPath.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
    $publishedFolderPath += [System.IO.Path]::DirectorySeparatorChar
}

$totalFound = $files.Count
$files = $files | Where-Object { 
    -not $_.FullName.StartsWith($publishedFolderPath, [System.StringComparison]::OrdinalIgnoreCase) 
}
$excludedCount = $totalFound - $files.Count

if ($excludedCount -gt 0) {
    Write-Verbose "Excluded $excludedCount files found in the Published folder."
}

if (-not $files) {
    Write-Host "No files found matching '$effectivePattern' under '$filesPath' (excluding Published folder)."
    exit 0
}

$successArray = @()
$failArray = @()
$projectSuccessArray = @()
$projectFailArray = @()
$baselineSuccessArray = @()
$baselineFailArray = @()

# Start logging to file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logsFolder = Join-Path $PSScriptRoot "Logs"
$logsFolder = Join-Path $logsFolder "Publish-MPPsToAltus_$timestamp"
# Create Logs folder if it doesn't exist
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}

# -----------------------------------------------------------------------------
# Build ordered file list with JSON pairing and date extraction
# Note: Find-JsonForMpp and Get-ProjectLastPublishedDate are in Common.ps1
# -----------------------------------------------------------------------------
$orderedFiles = @()
$pairingWarnings = @()
$maxDate = [DateTime]::MaxValue  # Used for EndOfQueue fallback

foreach ($mppFile in $files) {
    $jsonFile = Find-JsonForMpp -MppFile $mppFile
    $jsonPath = if ($jsonFile) { $jsonFile.FullName } else { $null }
    $lastPublishedDate = $null
    $dateSource = 'None'
    $warning = $null
    
    if ($jsonPath) {
        $lastPublishedDate = Get-ProjectLastPublishedDate -JsonPath $jsonPath
        if ($lastPublishedDate) {
            $dateSource = 'JSON'
        } else {
            $warning = "JSON file found but ProjectLastPublishedDate is missing or invalid: $jsonPath"
        }
    }
    
    # Apply fallback if no valid date from JSON
    if (-not $lastPublishedDate) {
        switch ($DateFallback) {
            'UseMppWriteTime' {
                $lastPublishedDate = $mppFile.LastWriteTime
                $dateSource = 'MppWriteTime'
            }
            'EndOfQueue' {
                $lastPublishedDate = $maxDate
                $dateSource = 'EndOfQueue'
            }
        }
        if (-not $jsonPath -and -not $warning) {
            $warning = "No matching JSON file found for: $($mppFile.FullName)"
        }
    }
    
    if ($warning) {
        $pairingWarnings += $warning
        Write-Warning $warning
    }
    
    $orderedFiles += [PSCustomObject]@{
        MppPath           = $mppFile.FullName
        MppWriteTime      = $mppFile.LastWriteTime
        JsonPath          = $jsonPath
        LastPublishedDate = $lastPublishedDate
        DateSource        = $dateSource
    }
}

# Sort by LastPublishedDate ascending (oldest first)
$orderedFiles = $orderedFiles | Sort-Object LastPublishedDate

# Ensure orderedFiles is an array even if it contains 0 or 1 item
if ($null -eq $orderedFiles) { $orderedFiles = @() }
if (-not ($orderedFiles -is [Array])) { $orderedFiles = @($orderedFiles) }

# Log the ordered list
$orderLogPath = Join-Path $logsFolder "publish-order.txt"
$orderLogContent = @()
$orderLogContent += "Publish Order Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$orderLogContent += "DateFallback Strategy: $DateFallback"
$orderLogContent += "DryRun: $DryRun"
$orderLogContent += "Total Files: $($orderedFiles.Count)"
$orderLogContent += ""
$orderLogContent += "Order | LastPublishedDate        | DateSource    | MPP Path"
$orderLogContent += "------+-------------------------+---------------+------------------------------------------"

$orderIndex = 1
foreach ($item in $orderedFiles) {
    $dateDisplay = if ($item.DateSource -eq 'EndOfQueue') { '(EndOfQueue)' } else { $item.LastPublishedDate.ToString('yyyy-MM-dd HH:mm:ss') }
    $orderLogContent += "{0,5} | {1,-23} | {2,-13} | {3}" -f $orderIndex, $dateDisplay, $item.DateSource, $item.MppPath
    $orderIndex++
}

$orderLogContent += ""
$orderLogContent += "Detailed Metadata:"
$orderLogContent += "=================="
foreach ($item in $orderedFiles) {
    $orderLogContent += ""
    $orderLogContent += "MPP:               $($item.MppPath)"
    $orderLogContent += "MPP WriteTime:     $($item.MppWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    $orderLogContent += "JSON:              $(if ($item.JsonPath) { $item.JsonPath } else { '(none)' })"
    $orderLogContent += "LastPublishedDate: $(if ($item.DateSource -eq 'EndOfQueue') { '(EndOfQueue)' } else { $item.LastPublishedDate.ToString('yyyy-MM-dd HH:mm:ss') })"
    $orderLogContent += "DateSource:        $($item.DateSource)"
}

if ($pairingWarnings.Count -gt 0) {
    $orderLogContent += ""
    $orderLogContent += "Warnings:"
    $orderLogContent += "========="
    foreach ($w in $pairingWarnings) {
        $orderLogContent += "  - $w"
    }
}

$orderLogContent | Out-File -FilePath $orderLogPath -Encoding utf8

# Export ordered metadata as CSV for easy analysis
$orderedFiles | Select-Object @{N = 'Order'; E = { [array]::IndexOf($orderedFiles, $_) + 1 } }, MppPath, MppWriteTime, JsonPath, LastPublishedDate, DateSource | 
Export-Csv -Path (Join-Path $logsFolder "publish-order.csv") -NoTypeInformation -Encoding utf8

Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Publish order determined. See: $orderLogPath" -ForegroundColor Cyan

# If DryRun, display the order and exit without publishing
if ($DryRun) {
    Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] DRY RUN - Files would be published in this order:" -ForegroundColor Yellow
    $orderIndex = 1
    foreach ($item in $orderedFiles) {
        $dateDisplay = if ($item.DateSource -eq 'EndOfQueue') { '(EndOfQueue)' } else { $item.LastPublishedDate.ToString('yyyy-MM-dd HH:mm:ss') }
        Write-Host ("  {0,3}. [{1}] {2}" -f $orderIndex, $dateDisplay, $item.MppPath)
        $orderIndex++
    }
    Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] DRY RUN complete. No files were published." -ForegroundColor Yellow
    exit 0
}

# Pre-flight network diagnostics (before starting publish)
$preFlight = Get-NetworkDiagnostics
if (-not $preFlight.NetworkAvailable) {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Error: Network connectivity test failed. Aborting publish run due to failed connectivity checks." -ForegroundColor Red
    exit 1
} else {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Network connectivity verified. Proceeding with publish." -ForegroundColor Green
}


# Create Published folder for successfully published files
$publishedFolder = Join-Path $filesPath "Published"
if (-not (Test-Path $publishedFolder)) {
    New-Item -Path $publishedFolder -ItemType Directory | Out-Null
    Write-Verbose "Created Published folder: $publishedFolder"
}
Open-WinProj

$automation = Get-AltusAutomationObject

# Initialize tracking variables for adaptive delays and COM restarts
$fileIndex = 0
$consecutiveFailures = 0
$consecutiveSuccesses = 0
$currentDelay = $DelayBetweenFiles
$currentRetryDelay = $RetryDelay
$filesSinceComRestart = 0
$comRestartCount = 0

# Adaptive delay configuration
$minDelay = [Math]::Max(1, [Math]::Floor($DelayBetweenFiles / 2))
$maxDelay = $DelayBetweenFiles * 4
$maxRetryDelay = $RetryDelay * 4

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Starting publish of $($orderedFiles.Count) file(s) with adaptive delays enabled." -ForegroundColor Cyan
Write-Host "  Initial delay between files: $currentDelay seconds" -ForegroundColor Gray
Write-Host "  Initial retry delay: $currentRetryDelay seconds" -ForegroundColor Gray
if ($ComRestartInterval -gt 0) {
    Write-Host "  COM restart interval: every $ComRestartInterval files" -ForegroundColor Gray
}

# Iterate through ordered files for publishing
foreach ($item in $orderedFiles) {
    $fileIndex++
    $filesSinceComRestart++
    $published = $false
    $attemptCount = 0
    $lastError = $null
    $result = $null
    $baselineResult = $null
    $comRestartedThisFile = $false
    
    $f = Get-Item $item.MppPath
    $projectName = Split-Path -Path $f.FullName -Leaf

    $safeProjectName = $projectName `
        -replace '[<>:"/\\|?*]', '_' `
        -replace '[,&()''`]', '_' `
        -replace '\s+', '_' `
        -replace '_+', '_' # Replace problematic characters with underscores for log file naming


    if ($safeProjectName.Length -gt 60) {
        Write-Warning "Project name '$projectName' is quite long. Truncating to ensure log file paths are manageable."
        $shortPrefix = $safeProjectName.Substring(0, 57) + "_"
    } else {
        $shortPrefix = $safeProjectName
    }

    $uniquePart = "_" + (Get-Date -Format "yyMMdd-HHmmss")

    $safeProjectName = $shortPrefix + $uniquePart

    if ($safeProjectName.Length -gt 100) {
        Write-Warning "Generated project name for logging is quite long (length: $($safeProjectName.Length))."
        $safeProjectName = $shortPrefix.Substring(0, 80) + $uniquePart.Substring(0, 12)
        Write-Warning "Truncated safe project name to: $safeProjectName (length: $($safeProjectName.Length)) to avoid excessively long file paths."
    }

    $exampleLogPath = Join-Path $logsFolder "$safeProjectName-publish-log.txt"
    $pathLength = $exampleLogPath.Length
    

    if ($pathLength -gt 260) {
        Write-Warning "Generated log file path exceeds typical Windows path length limits (length: $pathLength). This may cause issues on systems without long paths enabled."
    }

    # Proactive COM restart to prevent exhaustion
    if ($ComRestartInterval -gt 0 -and $filesSinceComRestart -ge $ComRestartInterval) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Proactive COM restart after $filesSinceComRestart files..." -ForegroundColor Magenta
        try {
            $automation = Restart-ComConnection -DelaySeconds $currentRetryDelay
            $filesSinceComRestart = 0
            $comRestartCount++
        } catch {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Failed to restart COM connection: $_" -ForegroundColor Red
            throw
        }
    }
    
    # Perform garbage collection periodically (every 5 files) and before retries
    if ($fileIndex % 5 -eq 0) {
        Write-Verbose "Performing garbage collection (file $fileIndex of $($orderedFiles.Count))..."
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
    
    while (-not $published -and $attemptCount -lt $RetryCount) {
        $attemptCount++
        
        if ($attemptCount -gt 1) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Retry attempt $attemptCount of $RetryCount for: $projectName" -ForegroundColor Yellow
            
            # Garbage collection before retry to free COM resources
            Write-Verbose "Performing garbage collection before retry..."
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
            
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Waiting $currentRetryDelay seconds before retry..." -ForegroundColor Gray
            Start-Sleep -Seconds $currentRetryDelay
        }
        
        try {
            # Ensure no stale projects are open before attempting to publish
            if ($attemptCount -gt 1 -or $fileIndex -gt 1) {
                try {
                    Close-AllOpenProjects
                } catch {
                    Write-Verbose "Pre-publish cleanup failed: $_"
                }
            }
                
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Publishing ($fileIndex/$($orderedFiles.Count)): $($f.FullName)"
            # Invoke-AltusPublishProject returns a tuple: (projectResult, baselineResult)
            $publishParams = @{
                projectPath     = $f.FullName
                automation      = $automation
                publishBaseline = $baselinePublishEnabled
            }
            if ($baselinePublishEnabled) {
                $publishParams['baselineNumber'] = $BaselineNumber
                $publishParams['isReportable'] = $IsReportable
            }
            $publishTuple = Invoke-AltusPublishProject @publishParams
            $result = $publishTuple.Item1
            $baselineResult = $publishTuple.Item2 # May be null
            
            if ($null -eq $result) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Publish failed: result is null." -ForegroundColor Red
                $lastError = "Invoke-AltusPublishProject returned null result."
                $published = $true  # Mark as handled - don't retry null results
                
                $failArray += [PSCustomObject]@{
                    Path      = $f.FullName
                    Error     = $lastError
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Attempts  = $attemptCount
                }
                $errorLogPath = Join-Path $logsFolder "$safeProjectName-null-result-error.txt"
                "Null result returned by Invoke-AltusPublishProject for $($f.FullName) at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $errorLogPath -Encoding utf8
                continue
            }

            # Save project logs (always, regardless of success/failure)
            $array = $result.Logs
            $resourceConfig = $result.ResourceConfig
            if ($null -ne $array) {
                $array | Out-File -FilePath "$logsFolder\$safeProjectName-publish-log.txt" -Encoding utf8
            }
            if ($null -ne $resourceConfig) {
                $resourceConfig | Out-File -FilePath "$logsFolder\$safeProjectName-resource-config.json" -Encoding utf8
            }

            # Save baseline logs if available (always, regardless of success/failure)
            if ($null -ne $baselineResult -and $null -ne $baselineResult.Logs) {
                $baselineResult.Logs | Out-File -FilePath "$logsFolder\$safeProjectName-baseline-log.txt" -Encoding utf8
            }

            # Check project publish result
            if ($result.Success) {
                $published = $true
                $consecutiveSuccesses++
                $consecutiveFailures = 0
                
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Project Publish succeeded." -ForegroundColor Green
                $successArray += $f.FullName
                $projectSuccessArray += $f.FullName
                
                # Move successfully published files to Published folder
                try {
                    $newMppPath = Join-Path $publishedFolder $f.Name
                    Move-Item -Path $f.FullName -Destination $newMppPath -Force
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Moved to Published folder: $newMppPath" -ForegroundColor Gray
                    
                    # Update reference to point to new location
                    $f = Get-Item $newMppPath
                    
                    # Move paired JSON metadata file if it exists
                    if ($item.JsonPath -and (Test-Path $item.JsonPath)) {
                        $jsonFileName = Split-Path -Path $item.JsonPath -Leaf
                        $newJsonPath = Join-Path $publishedFolder $jsonFileName
                        Move-Item -Path $item.JsonPath -Destination $newJsonPath -Force
                        Write-Verbose "Moved JSON metadata to Published folder: $newJsonPath"
                    }
                } catch {
                    Write-Warning "Failed to move published file to Published folder: $_"
                }
                
                # Adaptive delay: decrease delay after consecutive successes
                if ($consecutiveSuccesses -ge 3 -and $currentDelay -gt $minDelay) {
                    $currentDelay = [Math]::Max($minDelay, $currentDelay - 1)
                    Write-Verbose "Adaptive delay decreased to $currentDelay seconds after $consecutiveSuccesses consecutive successes."
                }
            } else {
                # Publish returned but with failure - don't retry, log the failure
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Publish failed." -ForegroundColor Yellow
                Write-Host "Error Message: $($result.Error)" -ForegroundColor Yellow
                $lastError = $result.Error
                $published = $true  # Mark as "handled" to exit retry loop - this is a logical failure, not a COM error
                
                $consecutiveFailures++
                $consecutiveSuccesses = 0
                
                $projectFailArray += [PSCustomObject]@{
                    Path      = $f.FullName
                    Error     = $result.Error
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Attempts  = $attemptCount
                }
                $failArray += [PSCustomObject]@{
                    Path      = $f.FullName
                    Error     = $result.Error
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Attempts  = $attemptCount
                }
                
                # Adaptive delay: increase delay after failures
                if ($currentDelay -lt $maxDelay) {
                    $currentDelay = [Math]::Min($maxDelay, $currentDelay + 2)
                    Write-Verbose "Adaptive delay increased to $currentDelay seconds after failure."
                }

                # If connectivity is down during publish failure, fail out immediately and clean up.
                if (-not (Test-NetworkConnectivity)) {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Error: Internet connectivity lost during publish. Closing Project and stopping run." -ForegroundColor Red
                    try { Close-AllOpenProjects } catch { Write-Verbose "Close-AllOpenProjects failed during network fail-out: $_" }
                    try { Close-WinProj } catch { Write-Verbose "Close-WinProj failed during network fail-out: $_" }
                    exit 1
                }

                # Treat common remote/connection error patterns as terminal (exit immediately on matched error text).
                if ($result.Error -match 'Object reference not set|remote connection has failed|connection.*failed') {
                    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Error: Add-in reported a remote/connection failure. Closing Project and stopping run." -ForegroundColor Red
                    try { Close-AllOpenProjects } catch { Write-Verbose "Close-AllOpenProjects failed during add-in fail-out: $_" }
                    try { Close-WinProj } catch { Write-Verbose "Close-WinProj failed during add-in fail-out: $_" }
                    exit 1
                }
            }
        } catch {
            $lastError = $_.ToString()
            $consecutiveFailures++
            $consecutiveSuccesses = 0
            
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Error during publish attempt $attemptCount`: $_" -ForegroundColor Red
            
            # Try to clean up any open projects before retry
            try {
                Close-AllOpenProjects
            } catch {
                Write-Verbose "Post-error cleanup failed: $_"
            }
            
            # Check for network connectivity issues on any error
            if (-not (Test-NetworkConnectivity)) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Error: Internet connectivity lost. Closing Project and stopping run." -ForegroundColor Red
                try { Close-AllOpenProjects } catch { Write-Verbose "Close-AllOpenProjects failed during exception network fail-out: $_" }
                try { Close-WinProj } catch { Write-Verbose "Close-WinProj failed during exception network fail-out: $_" }
                exit 1
            }
            
            # Check if this is an RPC/COM connection error
            if (Test-IsRpcError -ErrorMessage $lastError) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Detected RPC/COM connection error. Will restart COM before retry." -ForegroundColor Magenta
                
                if ($attemptCount -lt $RetryCount) {
                    try {
                        $automation = Restart-ComConnection -DelaySeconds ([Math]::Max($currentRetryDelay, 10)) -MaxAttempts 5
                        $filesSinceComRestart = 0
                        $comRestartCount++
                        $comRestartedThisFile = $true
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] COM connection restored. Retrying file..." -ForegroundColor Green
                        
                    } catch {
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Failed to restart COM connection: $_" -ForegroundColor Red
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Script cannot continue without COM connection. Exiting." -ForegroundColor Red
                        
                        # Save what we have before exiting
                        if ($successArray.Count -gt 0) {
                            $successArray | Out-File -FilePath "$logsFolder\successes.txt" -Encoding utf8
                        }
                        $failArray += [PSCustomObject]@{
                            Path      = $f.FullName
                            Error     = "COM connection lost and could not be restored: $lastError"
                            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                            Attempts  = $attemptCount
                        }
                        $failArray | Format-Table -AutoSize | Out-File -FilePath "$logsFolder\failures.txt" -Encoding utf8
                        $failArray | Export-Csv -Path "$logsFolder\failures.csv" -NoTypeInformation -Encoding utf8
                        
                        # Capture diagnostics at time of failure
                        $failureDiagnostics = Get-NetworkDiagnostics
                        @"
Network Diagnostics at COM Failure - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
==================================================================
Network Available: $($failureDiagnostics.NetworkAvailable)
DNS Resolution: $($failureDiagnostics.DnsResolution)
Localhost Available: $($failureDiagnostics.LoopbackAvailable)
Local Interfaces Up: $($failureDiagnostics.LocalNetworkFound)
Errors: $($failureDiagnostics.ErrorDetails.Count)
Details: $($failureDiagnostics.ErrorDetails -join ', ')
"@ | Out-File -FilePath "$logsFolder\network-diagnostics-at-failure.txt" -Encoding utf8
                        
                        Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Partial results saved. Processed $($fileIndex - 1) of $($orderedFiles.Count) files before failure." -ForegroundColor Yellow
                        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Succeeded: $($successArray.Count), Failed: $($failArray.Count)." -ForegroundColor Yellow
                        try { Close-AllOpenProjects } catch { Write-Verbose "Close-AllOpenProjects failed during COM fail-out: $_" }
                        try { Close-WinProj } catch { Write-Verbose "Close-WinProj failed during COM fail-out: $_" }
                        exit 1
                    }
                }
            }
            
            # Adaptive retry delay: increase retry delay after consecutive failures
            if ($consecutiveFailures -ge 2 -and $currentRetryDelay -lt $maxRetryDelay) {
                $currentRetryDelay = [Math]::Min($maxRetryDelay, $currentRetryDelay + 3)
                Write-Verbose "Adaptive retry delay increased to $currentRetryDelay seconds after $consecutiveFailures consecutive failures."
            }
            
            # Adaptive delay: increase delay between files after exceptions
            if ($currentDelay -lt $maxDelay) {
                $currentDelay = [Math]::Min($maxDelay, $currentDelay + 3)
                Write-Verbose "Adaptive delay increased to $currentDelay seconds after exception."
            }
            
            if ($attemptCount -ge $RetryCount) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Failed to publish after $RetryCount attempts: $($f.FullName)" -ForegroundColor Red
                $failArray += [PSCustomObject]@{
                    Path      = $f.FullName
                    Error     = $lastError
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Attempts  = $attemptCount
                }
            }
        }
    }
    
    # Check baseline publish result (if it exists) - AFTER all retry attempts
    if ($null -ne $baselineResult) {
        if ($baselineResult.Success) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Baseline publish succeeded." -ForegroundColor Green
            $baselineSuccessArray += $f.FullName
            $successArray += $f.FullName  # Also count baseline success towards overall success
        } else {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Baseline publish failed." -ForegroundColor Yellow
            Write-Host "Baseline Error Message: $($baselineResult.ErrorMessage)" -ForegroundColor Yellow
            $baselineFailArray += [PSCustomObject]@{
                Path      = $f.FullName
                Error     = $baselineResult.ErrorMessage
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            $failArray += [PSCustomObject]@{
                Path      = $f.FullName
                Error     = "Baseline publish failed: $($baselineResult.ErrorMessage)"
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Attempts  = 1
            }
        }
    } elseif ($baselinePublishEnabled) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $($f.FullName) - Baseline publish skipped (method not available or project publish failed)." -ForegroundColor Gray
    }
    
    # Write logs if we have a result (regardless of success/failure)
    if ($null -ne $result) {
        if ($null -ne $result.Logs) {
            $result.Logs | Out-File -FilePath "$logsFolder\$safeProjectName-publish-log.txt" -Encoding utf8
        } else {
            "No logs returned for $($f.FullName)" | Out-File -FilePath "$logsFolder\$safeProjectName-publish-log.txt" -Encoding utf8
        }
        if ($null -ne $result.ResourceConfig) {
            $result.ResourceConfig | Out-File -FilePath "$logsFolder\$safeProjectName-resource-config.json" -Encoding utf8
        } else {
            "No resource config returned for $($f.FullName)" | Out-File -FilePath "$logsFolder\$safeProjectName-resource-config.json" -Encoding utf8
        }
        if (-not [String]::IsNullOrEmpty($result.Error)) {
            if ($result.Success) {
                $result.Error | Out-File -FilePath "$logsFolder\$safeProjectName-errors.json" -Encoding utf8
                Write-Warning "Publish completed with errors. See: '$logsFolder\$safeProjectName-errors.json'";
            } else {
                $result.Error | Out-File -FilePath "$logsFolder\$safeProjectName-errors.txt" -Encoding utf8
            }
        }

        # Handle baseline errors
        if ($null -ne $baselineResult -and -not [String]::IsNullOrEmpty($baselineResult.ErrorMessage)) {
            if ($baselineResult.Success) {
                $baselineResult.ErrorMessage | Out-File -FilePath "$logsFolder\$safeProjectName-baseline-errors.json" -Encoding utf8
                Write-Warning "Baseline publish completed with errors. See: '$logsFolder\$safeProjectName-baseline-errors.json'"
            } else {
                $baselineResult.ErrorMessage | Out-File -FilePath "$logsFolder\$safeProjectName-baseline-errors.txt" -Encoding utf8
            }
        }
    }
    
    # Delay between files (skip after the last file)
    if ($fileIndex -lt $orderedFiles.Count) {
        Write-Verbose "Waiting $currentDelay seconds before next file..."
        Start-Sleep -Seconds $currentDelay
        
        # Perform periodic network health check every 10 files or if we've had failures
        if ($fileIndex % 10 -eq 0 -or $consecutiveFailures -ge 2) {
            $networkHealthCheck = Test-NetworkConnectivity
            if (-not $networkHealthCheck) {
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Network connectivity check between files FAILED at file $fileIndex/$($orderedFiles.Count)" -ForegroundColor Yellow
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Proceeding cautiously. Network may be unstable." -ForegroundColor Yellow
            } else {
                Write-Verbose "Network connectivity verified after file $fileIndex/$($orderedFiles.Count)."
            }
        }
    }
}

# Only write files if there's content
if ($successArray.Count -gt 0) {
    $successArray | Out-File -FilePath "$logsFolder\successes.txt" -Encoding utf8
}
if ($failArray.Count -gt 0) {
    $failArray | Format-Table -AutoSize | Out-File -FilePath "$logsFolder\failures.txt" -Encoding utf8
    $failArray | Export-Csv -Path "$logsFolder\failures.csv" -NoTypeInformation -Encoding utf8
}

# Final garbage collection
Write-Verbose "Performing final garbage collection..."
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

# Capture final network diagnostics
$postFlightDiagnostics = Get-NetworkDiagnostics

$filesWithRetries = ($failArray | Where-Object { $_.Attempts -gt 1 }).Count

Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Completed. Attempted to publish $($orderedFiles.Count) file(s). Succeeded: $($successArray.Count), Failed: $($failArray.Count)." -ForegroundColor Cyan
if ($filesWithRetries -gt 0) {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $filesWithRetries file(s) failed after multiple attempts." -ForegroundColor Yellow
}
if ($comRestartedThisFile -and $comRestartCount -gt 0) {
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] COM connection was restarted $comRestartCount time(s) during processing." -ForegroundColor Gray
}
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Final adaptive delay: $currentDelay seconds (started at $DelayBetweenFiles seconds)." -ForegroundColor Gray

# Display detailed summary
Write-Host "`n" -NoNewline
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "                              PUBLISH SUMMARY                                   " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Files Processed:    $($orderedFiles.Count)" -ForegroundColor White
Write-Host ""
Write-Host "  PROJECT PUBLISH:" -ForegroundColor White
Write-Host "    Succeeded:              $($projectSuccessArray.Count)" -ForegroundColor Green
Write-Host "    Failed:                 $($projectFailArray.Count)" -ForegroundColor $(if ($projectFailArray.Count -gt 0) { 'Red' } else { 'Green' })

if ($baselinePublishEnabled) {
    Write-Host ""
    Write-Host "  BASELINE PUBLISH (Baseline $BaselineNumber, Reportable: $IsReportable):" -ForegroundColor White
    Write-Host "    Succeeded:              $($baselineSuccessArray.Count)" -ForegroundColor Green
    Write-Host "    Failed:                 $($baselineFailArray.Count)" -ForegroundColor $(if ($baselineFailArray.Count -gt 0) { 'Red' } else { 'Green' })
}

Write-Host ""
Write-Host "  OVERALL:" -ForegroundColor White
Write-Host "    Succeeded:              $($successArray.Count)" -ForegroundColor Green
Write-Host "    Failed:                 $($failArray.Count)" -ForegroundColor $(if ($failArray.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host "    Post-Flight Available:  $($postFlightDiagnostics.NetworkAvailable)" -ForegroundColor $(if ($postFlightDiagnostics.NetworkAvailable) { 'Green' } else { 'Yellow' })
Write-Host ""
Write-Host "  Logs Folder:              $logsFolder" -ForegroundColor Gray
Write-Host "================================================================================" -ForegroundColor Cyan

# List failed projects if any
if ($projectFailArray.Count -gt 0) {
    Write-Host ""
    Write-Host "  FAILED PROJECT PUBLISHES:" -ForegroundColor Red
    foreach ($fail in $projectFailArray) {
        Write-Host "    - $($fail.Path)" -ForegroundColor Yellow
        Write-Host "      Error: $($fail.Error)" -ForegroundColor Gray
    }
}

if ($baselinePublishEnabled -and $baselineFailArray.Count -gt 0) {
    Write-Host ""
    Write-Host "  FAILED BASELINE PUBLISHES:" -ForegroundColor Red
    foreach ($fail in $baselineFailArray) {
        Write-Host "    - $($fail.Path)" -ForegroundColor Yellow
        Write-Host "      Error: $($fail.Error)" -ForegroundColor Gray
    }
}

Write-Host ""

Close-WinProj
