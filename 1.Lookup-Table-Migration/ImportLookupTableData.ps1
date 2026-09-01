<#
.SYNOPSIS
  Import lookup table data from CSV files into Dataverse.

.DESCRIPTION
  Imports lookup table data from CSV files (exported from Project Online) into Dataverse tables.
  Supports both flat lookup tables and hierarchical self-referential lookup tables.
  
  The script reads lookup table mapping configurations from LookupTablesToImport.ps1 and 
  processes each CSV file to create or update records in the corresponding Dataverse tables.

.PARAMETER D365Url
  The Dataverse environment URL (e.g., https://yourorg.crm.dynamics.com/)
  If not provided, you will be prompted to enter it.

.PARAMETER MigrationExePath
    Path to the separately approved Altus.DevOps.D365.DataMigration.exe tool.
    The executable is not bundled in this repository.

.PARAMETER Force
  When $true, allows updating existing records. When $false, only creates new records.
  Default: $true

.PARAMETER ParallelRequests
  Number of parallel requests to use when importing data.
  Default: 4

.PARAMETER EnableDisablingOfPlugins
  When $true, attempts to disable plugins during import for better performance.
  Requires elevated permissions. Default: $false

.EXAMPLE
    .\ImportLookupTableData.ps1 -D365Url "https://orgname.crm.dynamics.com/" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe"

.EXAMPLE
    .\ImportLookupTableData.ps1 -D365Url "https://orgname.crm.dynamics.com/" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe" -Force $false

.NOTES
    - Historical mapping reference: LookupTablesToImport.ps1
  - CSV files must be in the LookupTables folder (or path specified in LookupTablesToImport.ps1)
  - Target Dataverse tables and fields must already exist
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$D365Url,

    [Parameter(Mandatory = $true)]
    [string]$MigrationExePath,

    [Parameter(Mandatory = $false)]
    [bool]$Force = $true,

    [Parameter(Mandatory = $false)]
    [int]$ParallelRequests = 4,

    [Parameter(Mandatory = $false)]
    [bool]$EnableDisablingOfPlugins = $false
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# ==============================================================================
# INITIALIZATION
# ==============================================================================

$OutputFolder = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "DataLoaderLookupTables")

Write-Host ""
Write-Host "Lookup Table Import Parameters:" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
# PREREQUISITES
# ==============================================================================

Write-Host "Checking prerequisites..." -ForegroundColor Cyan

# Check PowerShell version (minimum 7.4)
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -lt 7) {
    # Check if PowerShell 7 is installed
    $pwshPath = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Path
    
    if (-not $pwshPath) {
        Write-Host "`nERROR: PowerShell 7 is not installed." -ForegroundColor Red
        Write-Host "Current version: $($psVersion.ToString())" -ForegroundColor Yellow
        Write-Host "Please download and install PowerShell 7.4 or later from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
        Write-Host "`nExiting..." -ForegroundColor Red
        exit 1
    }
    
    # Relaunch in PowerShell 7 with same parameters
    Write-Host "Relaunching script in PowerShell 7..." -ForegroundColor Yellow
    
    # Build command string to pass to pwsh
    $scriptPath = $MyInvocation.MyCommand.Path
    $cmdParts = @("& '$scriptPath'")
    if ($D365Url) { $cmdParts += "-D365Url '$D365Url'" }
    if ($MigrationExePath) { $cmdParts += "-MigrationExePath '$MigrationExePath'" }
    $cmdParts += "-Force `$$Force"
    if ($ParallelRequests) { $cmdParts += "-ParallelRequests $ParallelRequests" }
    $cmdParts += "-EnableDisablingOfPlugins `$$EnableDisablingOfPlugins"
    
    $command = $cmdParts -join ' '
    & pwsh.exe -Command $command
    exit
} elseif ($psVersion.Major -lt 7 -or ($psVersion.Major -eq 7 -and $psVersion.Minor -lt 4)) {
    Write-Host "ERROR: PowerShell 7.4 or later is required." -ForegroundColor Red
    Write-Host "Current version: $($psVersion.ToString())" -ForegroundColor Yellow
    Write-Host "Please download and install the latest PowerShell from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
    Write-Host "Exiting..." -ForegroundColor Red
    exit 1
}
Write-Host "  OK PowerShell version $($psVersion.ToString()) meets minimum requirement of 7.4" -ForegroundColor Green

# Create output folder
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
    Write-Host ""
}

# Tracking
$statsLookupTables = 0
$statsErrors = 0
$hasErrors = $false
$statsSuccess = 0
$successCount = 0
$failureCount = 0
$results = @()

# Start logging to file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logsFolder = Join-Path $PSScriptRoot "Logs"
# Create Logs folder if it doesn't exist
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}
$logFile = Join-Path $logsFolder "ImportLookupTableData_$timestamp.txt"
Start-Transcript -Path $logFile -Append

# Import required mappings
. $PSScriptRoot\LookupTablesToImport.ps1
# Import DataLoader Common functions
. "$PSScriptRoot\..\4.SharePoint-Migration\CommonFunctions.ps1"
# Import DataLoader specific functions for lookup tables (includes Import-LookupTableData function)
. $PSScriptRoot\LookupTableDataLoaderOperations.ps1

Write-Host "       " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host ""
Write-Host " Ʌ̲LTUS " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host ""
Write-Host "       " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host ""
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Project Online Lookup Table Value Import Script" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

# Prompt for Dataverse environment URL if not provided
if (-not $D365Url) {
    $D365Url = Read-Host -Prompt "Enter the Dataverse environment URL (e.g. https://yourorg.crm.dynamics.com/)"
}

# Ensure the URL ends with a trailing slash
if (-not $D365Url.EndsWith('/')) {
    $D365Url += '/'
}

Write-Host ""

# ==============================================================================
# CREATE OUTPUT
# ==============================================================================

$dataXml = New-CmtDataXml

if (-not (Test-Path $OutputFolder)) { 
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null 
}

# ==============================================================================
# PROCESS LOOKUP TABLE MAPPINGS
# ==============================================================================
#iterate through each lookup table defined at the top of the script and generate a corresponding dataloader import file
#Data loader file will preserve the id values from Project Online to allow for easier mapping during migration
foreach ($lookupTableMapping in $LookupTableMappings) {
    $sourceFile = $lookupTableMapping.SourceFile
    $mapAsHierarchy = $lookupTableMapping.MapAsHierarchy
    $dataverseTable = $lookupTableMapping.DataverseTable
    Write-Host "Generating Dataloader import file for lookup table from file: $sourceFile" -ForegroundColor Cyan

    #Generate Dataloader XML from CSV
    try {
        $sourceFilePath = Join-Path $LookupTableFileLocation $sourceFile
        $lookupData = Import-Csv -Path $sourceFilePath -Encoding UTF8
        if (-not $lookupData -or $lookupData.Count -eq 0) {
            Write-Host "No data found in file: $sourceFile" -ForegroundColor Yellow
            continue
        }

        # Load code mask to determine separators
        $codeMaskFile = $sourceFile -replace '^LookupTable_', 'CodeMask_'
        $codeMaskPath = Join-Path $LookupTableFileLocation $codeMaskFile
        $separators = @()
        
        if (Test-Path $codeMaskPath) {
            $codeMask = Import-Csv -Path $codeMaskPath -Encoding UTF8
            $separators = $codeMask | ForEach-Object { $_.Separator }
            Write-Host "  Code mask loaded: $($separators.Count) levels with separators: $($separators -join ', ')" -ForegroundColor Gray
        } else {
            Write-Host "  No code mask found, using default separators" -ForegroundColor Yellow
        }

        # ==============================================================================
        # ADD LOOKUP DATA TO XML    
        # ==============================================================================

        foreach ($entry in $LookupData) {
            $attrs = Convert-LookupValueToEntity -Item $entry -LookupTableMapping $lookupTableMapping -LookupData $lookupData -CodeMaskSeparators $separators
            # Create deterministic GUID from SharePoint List ID and Item ID to ensure consistency across multiple exports
            $attrs["_recordId"] = $entry.Id;
        
            Add-CmtEntityRecord -Doc $dataXml -EntityLogicalName $dataverseTable.TableLogicalName -Attributes $attrs
        }
    } catch {
        Write-Host "Error generating dataloader XML from $sourceFile : $($_.Exception.Message)" -ForegroundColor Red
        $statsErrors++
        $hasErrors = $true
        continue
    }
    $statsLookupTables++
    Write-Host "  OK: Generated Dataloader XML from $sourceFile" -ForegroundColor Green
}

# ==============================================================================
# SAVE LOOKUP DATA LOADER XML
# ==============================================================================

$dataPath = Join-Path $OutputFolder "Data.xml"
$dataXml.Save($dataPath)
Write-Host "  Data.xml saved to: $dataPath" -ForegroundColor Green

# ==============================================================================
# IMPORT LOOKUP DATA TO DATAVERSE
# ==============================================================================
# TO DO: Use data loader to actually import the data
Write-Host ("-" * 80) -ForegroundColor Gray
Write-Host "Processing..." -ForegroundColor Cyan
Write-Host "  Data File: $dataPath"
Write-Host ""

# Build arguments for migration tool
$arguments = @(
    "-UseInteractive", "true",
    "-D365Url", $D365Url,
    #    "-SchemaPath", "`"$SchemaPath`"",
    "-DataPath", $dataPath,
    "-Force", $Force.ToString().ToLower(),
    "-ParallelRequests", $ParallelRequests.ToString(),
    "-EnableDisablingOfPlugins", $EnableDisablingOfPlugins.ToString().ToLower()
)

# Execute migration
$startTime = Get-Date
Write-Host "Starting import at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray

try {
    # Call the executable and capture output
    $output = & $MigrationExePath @arguments 2>&1
    $exitCode = $LASTEXITCODE
    
    # Write output to host so it gets captured by transcript
    $output | ForEach-Object { Write-Host $_ }
    
    # Check if output contains error indicators (even if exit code is 0)
    $outputText = $output | Out-String
    $hasDataLoaderErrors = $false
    $dataLoaderFailedCount = 0
    $dataLoaderSuccessCount = 0
    
    # Extract successful record count from the summary (supports old and new formats)
    if ($outputText -match '(?im)^\s*-\s*RecordsProcessedSuccessfully:\s*([\d,]+)\s*$') {
        $dataLoaderSuccessCount = [int](($matches[1] -replace ',', '').Trim())
    } elseif ($outputText -match 'Successful:\s+([\d,]+)\s+\([\d.]+%\)') {
        $dataLoaderSuccessCount = [int](($matches[1] -replace ',', '').Trim())
    }

    # Extract failed record count (supports old and new formats)
    if ($outputText -match '(?im)^\s*-\s*RecordsFailed:\s*([\d,]+)\s*$') {
        $dataLoaderFailedCount = [int](($matches[1] -replace ',', '').Trim())
    } elseif ($outputText -match 'Failed:\s+([\d,]+)\s+\([\d.]+%\)') {
        $dataLoaderFailedCount = [int](($matches[1] -replace ',', '').Trim())
    }
    
    # Check for error report heading and extract failed count
    if ($outputText -match 'IMPORT ERROR REPORT') {
        $hasDataLoaderErrors = $true
        
        if ($dataLoaderFailedCount -gt 0) {
            Write-Host "WARNING: Data loader reported $dataLoaderFailedCount failed records" -ForegroundColor Yellow
        } else {
            Write-Host "WARNING: Data loader reported errors during import (ERROR REPORT found)" -ForegroundColor Yellow
        }
    }
    # Check for failed record count even without error report
    elseif ($dataLoaderFailedCount -gt 0) {
        $hasDataLoaderErrors = $true
        Write-Host "WARNING: Data loader reported $dataLoaderFailedCount failed records" -ForegroundColor Yellow
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    if ($exitCode -eq 0 -and -not $hasDataLoaderErrors) {
        Write-Host "OK: SUCCESS - $dataPath imported successfully" -ForegroundColor Green
        Write-Host "  Duration: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        $successCount++
        $statsSuccess += $dataLoaderSuccessCount
        
        $results += [PSCustomObject]@{
            Status            = "Success"
            ExitCode          = $exitCode
            Duration          = $duration
            DataPath          = $dataPath
            SuccessfulRecords = $dataLoaderSuccessCount
        }
    } else {
        $errorReason = if ($exitCode -ne 0) { "exit code $exitCode" } else { "$dataLoaderFailedCount data loader errors" }
        Write-Host "FAILED: $dataPath import failed with $errorReason" -ForegroundColor Red
        Write-Host "  Duration: $($duration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
        $failureCount++
        $statsErrors += if ($dataLoaderFailedCount -gt 0) { $dataLoaderFailedCount } else { 1 }
        $statsSuccess += $dataLoaderSuccessCount
        $hasErrors = $true
        
        $results += [PSCustomObject]@{
            Status            = "Failed"
            ExitCode          = $exitCode
            Duration          = $duration
            DataPath          = $dataPath
            FailedRecords     = $dataLoaderFailedCount
            SuccessfulRecords = $dataLoaderSuccessCount
        }
    }
} catch {
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "ERROR: $displayName - $($_.Exception.Message)" -ForegroundColor Red
    $failureCount++
    
    $results += [PSCustomObject]@{
        Status   = "Error"
        ExitCode = -1
        Duration = $duration
        DataPath = $dataPath
        Error    = $_.Exception.Message
    }

    $hasErrors = $true
    $statsErrors++
}

Write-Host ""

Write-Host "Import Summary:" -ForegroundColor Cyan
Write-Host "---------------" -ForegroundColor Cyan
Write-Host "  Total Lookup Table Files Processed: $($statsLookupTables)" -ForegroundColor Gray
Write-Host "  Successful Imports: $($successCount)" -ForegroundColor Green
Write-Host "  Successful Records: $($statsSuccess)" -ForegroundColor Green
$failColor = if ($failureCount -eq 0) { 'Green' } else { 'Red' }
Write-Host "  Failed Imports: $($failureCount)" -ForegroundColor $failColor
if ($hasErrors) {
    Write-Host "  Total Errors: $($statsErrors)" -ForegroundColor Red
} else {
    Write-Host "  Total Errors: $($statsErrors)" -ForegroundColor Green
}
Write-Host ""
Write-Host "Lookup table data import completed." -ForegroundColor Green
