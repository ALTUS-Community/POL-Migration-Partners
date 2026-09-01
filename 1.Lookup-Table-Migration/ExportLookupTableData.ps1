<#
.SYNOPSIS
  Export lookup tables and values from a Project Online PWA instance.

.DESCRIPTION
  Exports all lookup tables and their entries from a Project Online PWA instance.

.PARAMETER PwaUrl
  The PWA site collection URL (e.g., https://tenant.sharepoint.com/sites/PWA)

.PARAMETER ClientId
    Optional Entra ID Client ID for authentication. Configure it for the approved working tool.

.PARAMETER OutputFolder
  Folder to write exported documents and manifests
#>

param(
    [Parameter(Mandatory = $false)]
    [string] $PwaUrl = "https://contoso.sharepoint.com/sites/pwa",

    [Parameter(Mandatory = $false)]
    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string] $AzureEnvironment = 'Production',

    [Parameter(Mandatory = $false)]
    [string] $ClientId = "YOUR-ENTRA-CLIENT-ID",

    [Parameter(Mandatory = $false)]
    [string] $OutputFolder = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "LookupTables")
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

<#
.SYNOPSIS
    Exports code mask configuration for a lookup table.

.DESCRIPTION
    Retrieves and exports the code mask configuration (hierarchy levels, separators, etc.)
    for a specified Project Server lookup table.

.PARAMETER LookupTableId
    The GUID of the lookup table.

.PARAMETER LookupTableName
    The name of the lookup table (for display purposes).

.PARAMETER OutputPath
    The file path where the code mask CSV will be saved.
#>
function Export-CodeMaskData {
    param (
        [string]$LookupTableId,
        [string]$LookupTableName,
        [string]$OutputPath
    )

    try {
        # Get lookup table with expanded Masks collection
        $lookupTableDetailUrl = "/_api/ProjectServer/LookupTables('$LookupTableId')?`$expand=Masks"
        $lookupTableDetail = Invoke-PnPSPRestMethod -Url $lookupTableDetailUrl -Method Get
        
        if (-not $lookupTableDetail.Masks -or $lookupTableDetail.Masks.Count -eq 0) {
            Write-Host "  No code mask defined for: $LookupTableName" -ForegroundColor Gray
            return
        }

        # Convert masks to export format
        $exportData = @()
        $defaultLevel = 0;
        foreach ($mask in $lookupTableDetail.Masks) {
            # Convert MaskType enum to readable string
            $maskTypeString = switch ($mask.MaskType) {
                0 { "Numbers" }
                1 { "UpperCaseLetters" }
                2 { "LowerCaseLetters" }
                3 { "Characters" }
                default { "Unknown" }
            }
            
            $exportData += [PSCustomObject]@{
                Level     = if ($null -ne $mask.Level) { $mask.Level } else { $defaultLevel++ }
                Length    = $mask.Length
                MaskType  = $maskTypeString
                Separator = $mask.Separator
            }
        }

        # Export to CSV
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported code mask ($($exportData.Count) levels) to $OutputPath" -ForegroundColor Green
        
    } catch {
        Write-Host "  Error exporting code mask for $LookupTableName - $($_.Exception.Message)" -ForegroundColor Red
    }
}

<#
.SYNOPSIS
    Exports lookup table entries to a CSV file.

.DESCRIPTION
    Retrieves all entries from a specified Project Server lookup table via REST API
    and exports them to a CSV file with key metadata including InternalName, FullValue,
    Description, SortIndex, and Id.

.PARAMETER LookupTable
    The lookup table object to export.

.PARAMETER OutputPath
    The file path where the CSV export will be saved.
#>
function Export-LookupTableData {
    param (
        [object]$LookupTable,
        [string]$OutputPath
    )

    try {
        # Get the lookup table ID
        $lookupTableId = $LookupTable.Id
        
        # Get lookup table entries via REST API
        $lookupEntriesUrl = "/_api/ProjectServer/LookupTables('$lookupTableId')/Entries"
        $lookupEntries = Invoke-PnPSPRestMethod -Url $lookupEntriesUrl -Method Get
        
        if (-not $lookupEntries.value -or $lookupEntries.value.Count -eq 0) {
            Write-Host "No entries found in lookup table: $($LookupTable.Name)" -ForegroundColor Yellow
            return
        }

        # Convert to export format
        $exportData = @()
        foreach ($entry in $lookupEntries.value) {
            $exportData += [PSCustomObject]@{
                InternalName = $entry.InternalName
                FullValue    = $entry.FullValue
                Description  = $entry.Description
                SortIndex    = $entry.SortIndex
                Id           = $entry.Id
            }
        }

        # Export to CSV
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported $($exportData.Count) entries from $($LookupTable.Name) to $OutputPath" -ForegroundColor Green
        
    } catch {
        Write-Host "Error exporting $($LookupTable.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

<#
.SYNOPSIS
    Checks if an exception message indicates an access denied error.

.PARAMETER ExceptionMessage
    The exception message string to check.
#>
function Test-AccessDeniedError {
    param([string]$ExceptionMessage)
    
    if ($ExceptionMessage -match "Access.*denied|Unauthorized|403") {
        return $true
    }
    return $false
}

# ==============================================================================
# PREREQUISITES
# ==============================================================================

# Check Prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Cyan

# Check PowerShell version (minimum 7.4)
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -lt 7) {
    # Check if PowerShell 7 is installed
    $pwshPath = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Path
    
    if (-not $pwshPath) {
        Write-Host ""
        Write-Host "ERROR: PowerShell 7 is not installed." -ForegroundColor Red
        Write-Host "Current version: $($psVersion.ToString())" -ForegroundColor Yellow
        Write-Host "Please download and install PowerShell 7.4 or later from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Exiting..." -ForegroundColor Red
        exit 1
    }
    
    # Relaunch in PowerShell 7 with same parameters
    Write-Host "Relaunching script in PowerShell 7..." -ForegroundColor Yellow
    
    # Build parameter list to pass to pwsh
    $paramList = @()
    if ($PwaUrl) { $paramList += "-PwaUrl", $PwaUrl }
    if ($AzureEnvironment) { $paramList += "-AzureEnvironment", $AzureEnvironment }
    if ($ClientId) { $paramList += "-ClientId", $ClientId }
    if ($OutputFolder) { $paramList += "-OutputFolder", $OutputFolder }
    
    & pwsh.exe -File $MyInvocation.MyCommand.Path @paramList
    exit
} elseif ($psVersion.Major -lt 7 -or ($psVersion.Major -eq 7 -and $psVersion.Minor -lt 4)) {
    Write-Host ""
    Write-Host "ERROR: PowerShell 7.4 or later is required." -ForegroundColor Red
    Write-Host "Current version: $($psVersion.ToString())" -ForegroundColor Yellow
    Write-Host "Please download and install the latest PowerShell from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Exiting..." -ForegroundColor Red
    exit 1
}
Write-Host "  [OK] PowerShell version $($psVersion.ToString()) meets minimum requirement of 7.4" -ForegroundColor Green


if (-not (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue)) {
    Write-Host "PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "Install with: Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

Write-Host "PnP.PowerShell module is installed"
Write-Host ""

# ==============================================================================
# INITIALIZATION
# ==============================================================================
# Start logging to file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logsFolder = Join-Path $PSScriptRoot "Logs"
# Create Logs folder if it doesn't exist
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}
$logFile = Join-Path $logsFolder "ExportLookupTableData_$timestamp.txt"
Start-Transcript -Path $logFile -Append

Write-Host ""
Write-Host "Lookup Table Export Parameters:" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan
Write-Host "PwaUrl:           $PwaUrl" -ForegroundColor Gray
Write-Host "AzureEnvironment: $AzureEnvironment" -ForegroundColor Gray
Write-Host "ClientId:         $ClientId" -ForegroundColor Gray
Write-Host "OutputFolder:     $OutputFolder" -ForegroundColor Gray
Write-Host ""
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
$scriptExitCode = 0

# ==============================================================================
# GLOBAL TRY-FINALLY TO GUARANTEE LOG SAVE
# ==============================================================================
try {

    # ==============================================================================
    # CONNECT TO PWA
    # ==============================================================================

    Write-Host "Connecting to SharePoint (interactive)..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $PwaUrl -ClientId $ClientId -AzureEnvironment $AzureEnvironment -Interactive -ErrorAction Stop
        Write-Host "Connected to: $PwaUrl"
        Write-Host ""
    } catch {
        Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
        
        # Check for access denied errors
        if (Test-AccessDeniedError -ExceptionMessage $_.Exception.Message) {
            Write-Host "" -ForegroundColor Red
            Write-Host "Please verify that you have Site Collection Administrator permissions" -ForegroundColor Red
            Write-Host "for this site collection: $PwaUrl" -ForegroundColor Red
            Write-Host "" -ForegroundColor Red
        }
        $hasErrors = $true
        $statsErrors++
        $scriptExitCode = 1
        throw
    }

    # ==============================================================================
    # PROCESS EACH LOOKUP TABLE
    # ==============================================================================
    # Get all lookup tables directly using Project Server REST API
    Write-Host "Retrieving all lookup tables from Project Online..." -ForegroundColor Cyan

    try {
        $lookupTablesUrl = "/_api/ProjectServer/LookupTables"
        $lookupTablesResponse = Invoke-PnPSPRestMethod -Url $lookupTablesUrl -Method Get
        $lookupTables = $lookupTablesResponse.value
        
        Write-Host "Found $($lookupTables.Count) lookup tables" -ForegroundColor Cyan
        
        # Export each lookup table
        foreach ($lookupTable in $lookupTables) {
            $sanitizedName = $lookupTable.Name -replace '[\\/:*?"<>|]', '_'
            $outputPath = Join-Path $OutputFolder "LookupTable_$sanitizedName.csv"
            $codeMaskPath = Join-Path $OutputFolder "CodeMask_$sanitizedName.csv"
            
            Write-Host "Exporting lookup table: $($lookupTable.Name)" -ForegroundColor Yellow
            Export-LookupTableData -LookupTable $lookupTable -OutputPath $outputPath
            Export-CodeMaskData -LookupTableId $lookupTable.Id -LookupTableName $lookupTable.Name -OutputPath $codeMaskPath
            $statsLookupTables++
        }
        
    } catch {
        $hasErrors = $true
        $statsErrors++
        if ($scriptExitCode -eq 0) {
            $scriptExitCode = 2
        }
        Write-Host "Error accessing Project Server REST API: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Make sure you're connected to a Project Online PWA site." -ForegroundColor Red
    }

    Write-Host "All lookup table exports completed!" -ForegroundColor Green

} catch {
    if ($scriptExitCode -eq 0) {
        $scriptExitCode = 1
    }
    Write-Host "Unhandled error: $($_.Exception.Message)" -ForegroundColor Red
} finally {   
    # ==============================================================================
    # SUMMARY
    # ==============================================================================

    Write-Host ""
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "Lookup Table Export Complete!" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Lookup Tables processed:  $statsLookupTables" -ForegroundColor White
    if ($statsErrors -gt 0) {
        Write-Host "Errors encountered:  $statsErrors" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Output location: $OutputFolder" -ForegroundColor Green
    Write-Host ""

    if ($scriptExitCode -eq 1) {
        Write-Host "Script failed. Review log file for details." -ForegroundColor Red
    } elseif ($scriptExitCode -eq 2 -or $hasErrors) {
        if ($scriptExitCode -eq 0) {
            $scriptExitCode = 2
        }
        Write-Host "Script completed with errors. Review log file for details." -ForegroundColor Yellow
    } else {
        Write-Host "Script completed successfully." -ForegroundColor Green
    }

    # Disconnect the session
    try {
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Disconnect-PnPOnline -ErrorAction Stop
        }
    } catch {
        Write-Host "Warning: Failed to disconnect PnP session: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Stop logging
    try {
        Stop-Transcript | Out-Null
    } catch {
    }
    Write-Host ""
    Write-Host "Log file saved to: $logFile"
}

exit $scriptExitCode

