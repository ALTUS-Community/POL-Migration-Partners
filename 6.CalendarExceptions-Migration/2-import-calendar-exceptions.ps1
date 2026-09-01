<#
.SYNOPSIS
    Imports the calendar-exceptions Data.xml into Dynamics 365 using the
    Altus.DevOps.D365.DataMigration console app.

.DESCRIPTION
    Thin wrapper over Altus.DevOps.D365.DataMigration.exe, following the same
    argument pattern as 4.SharePoint-Migration/2-import-lists.ps1. Unlike that
    script, this one does not discover or merge multiple per-project Data.xml
    files - 1-export-calendar-exceptions.ps1 already produces exactly one
    consolidated Data.xml for the whole batch, so -DataFolder simply locates it.

.PARAMETER D365Url
    The URL of your Dynamics 365 environment (e.g., https://orgname.crm.dynamics.com)

.PARAMETER DataFolder
    Folder containing Data.xml (the -OutputFolder used by 1-export-calendar-exceptions.ps1)

.PARAMETER MigrationExePath
    Path to the separately approved Altus.DevOps.D365.DataMigration.exe executable.
    The executable is not bundled in this repository.

.PARAMETER Force
    If true, updates existing records. Default is true (matches 2-import-lists.ps1)

.PARAMETER ParallelRequests
    Number of parallel requests for faster import. Default is 4

.PARAMETER EnableDisablingOfPlugins
    If true, bypasses plugins during migration. Default is false

.EXAMPLE
    .\2-import-calendar-exceptions.ps1 -D365Url "https://orgname.crm.dynamics.com" -DataFolder ".\Output\CalendarExceptions" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe"

.EXAMPLE
    .\2-import-calendar-exceptions.ps1 -D365Url "https://orgname.crm.dynamics.com" -DataFolder ".\Output\CalendarExceptions" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe" -Force $true -ParallelRequests 8
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $D365Url,

    [Parameter(Mandatory = $true)]
    [string] $DataFolder,

    [Parameter(Mandatory = $true)]
    [string] $MigrationExePath,

    [Parameter(Mandatory = $false)]
    [bool] $Force = $true,

    [Parameter(Mandatory = $false)]
    [int] $ParallelRequests = 4,

    [Parameter(Mandatory = $false)]
    [bool] $EnableDisablingOfPlugins = $false
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# ==============================================================================
# BOOTSTRAP
# ==============================================================================

if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$scriptFolder = $PSScriptRoot

# Import PS5-safe relaunch helper and elevate to pwsh when required.
$relaunchHelperPath = Join-Path $scriptFolder "..\4.SharePoint-Migration\Compat\Relaunch.PS5Safe.ps1"
if (-not (Test-Path $relaunchHelperPath)) {
    Write-Error "Relaunch helper script not found: $relaunchHelperPath"
    $global:LASTEXITCODE = 1
    return
}
. $relaunchHelperPath
Invoke-RelaunchInPwshIfNeeded -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters

# Import common helpers (Resolve-RelativePath)
$commonHelpersPath = Join-Path $scriptFolder "..\4.SharePoint-Migration\CommonFunctions.ps1"
if (-not (Test-Path $commonHelpersPath)) {
    Write-Error "Common helpers script not found: $commonHelpersPath"
    $global:LASTEXITCODE = 1
    return
}
. $commonHelpersPath

# ==============================================================================
# PATH RESOLUTION
# ==============================================================================

$DataFolder = Resolve-RelativePath -Path $DataFolder -BasePath $scriptFolder
$MigrationExePath = Resolve-RelativePath -Path $MigrationExePath -BasePath $scriptFolder

if (-not (Test-Path $DataFolder -PathType Container)) {
    Write-Error "DataFolder not found: $DataFolder"
    $global:LASTEXITCODE = 1
    return
}

$dataPath = Join-Path $DataFolder "Data.xml"
if (-not (Test-Path $dataPath -PathType Leaf)) {
    Write-Error "Data.xml not found in DataFolder: $dataPath"
    $global:LASTEXITCODE = 1
    return
}

if (-not (Test-Path $MigrationExePath)) {
    Write-Error "Migration executable not found: $MigrationExePath"
    $global:LASTEXITCODE = 1
    return
}

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Dynamics 365 Data Import - Calendar Exceptions" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  D365 URL:              $D365Url"
Write-Host "  Data File:             $dataPath"
Write-Host "  Migration Tool:        $MigrationExePath"
Write-Host "  Force Update:          $Force"
Write-Host "  Parallel Requests:     $ParallelRequests"
Write-Host "  Disable Plugins:       $EnableDisablingOfPlugins"
Write-Host ""

# ==============================================================================
# DATAMIGRATION.EXE INVOCATION
# ==============================================================================
# Same argument pattern as 4.SharePoint-Migration/2-import-lists.ps1, without its
# multi-project discovery/merge logic - there is only ever one Data.xml here.

$arguments = @(
    "-UseInteractive", "true",
    "-D365Url",                  "`"$D365Url`"",
    "-DataPath",                  "`"$dataPath`"",
    "-Force",                     "`"$($Force.ToString().ToLower())`"",
    "-ParallelRequests",          $ParallelRequests.ToString(),
    "-EnableDisablingOfPlugins",  "`"$($EnableDisablingOfPlugins.ToString().ToLower())`""
)

$startTime = Get-Date
Write-Host "Starting import at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray

try {
    $process = Start-Process -FilePath $MigrationExePath `
        -ArgumentList $arguments `
        -Wait `
        -PassThru `
        -NoNewWindow

    $endTime = Get-Date
    $duration = $endTime - $startTime

    if ($process.ExitCode -eq 0) {
        Write-Host "[SUCCESS] Calendar exceptions imported successfully" -ForegroundColor Green
        Write-Host "  Duration: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        $global:LASTEXITCODE = 0
    } else {
        Write-Host "[FAILED] Import failed with exit code $($process.ExitCode)" -ForegroundColor Red
        Write-Host "  Duration: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        $global:LASTEXITCODE = 1
    }
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    $global:LASTEXITCODE = 1
}
