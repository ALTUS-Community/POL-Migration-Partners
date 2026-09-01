<#
.SYNOPSIS
Main orchestration script for migrating Project Online Task field data to Altus (Dataverse).

.DESCRIPTION
This script coordinates the migration of task level fields (in particular, custom fields) from Project Online 
to Altus Project Management. It handles:
- Connection to the Dataverse environment
- Loading lookup and choice field data
- Updating tasks to correctly reflect custom field values

The script validates prerequisites (PowerShell version, Azure CLI) and provides 
interactive prompts for configuration if parameters are not provided.

.PARAMETER EnvironmentUrl
The Dataverse environment URL (e.g., https://yourorg.crm.dynamics.com/)
If not provided, you will be prompted to enter it interactively.

.PARAMETER ExecutionMode
Controls whether to actually execute the import or run in test mode:
- 'E' or 'Execute': Creates/updates records in Dataverse
- 'W' or 'WhatIf': Test mode - shows what would happen without making changes
If not provided, you will be prompted to select the mode interactively.

.PARAMETER ForceLogin
When specified, forces a fresh Azure CLI login prompt even if valid cached credentials exist.
Use this when you need to authenticate as a different user with higher permissions in the
target environment (e.g. an Altus Admin rather than a read-only account).

.EXAMPLE
.\AltusTaskFieldMigration.ps1

Runs interactively, prompting for all parameters.

.EXAMPLE
.\AltusTaskFieldMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode WhatIf

Runs in test mode to update task fields without making actual changes.

.EXAMPLE
.\AltusTaskFieldMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode Execute

Updates task fields in Dataverse.

.EXAMPLE
.\AltusTaskFieldMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode Execute -ForceLogin

Forces a fresh login prompt before updating task fields, allowing you to authenticate as a
different user (e.g. one with Altus Admin permissions).

.NOTES
Prerequisites:
- PowerShell 7.4 or later
- Azure CLI (https://aka.ms/installazurecliwindows) — used for authentication to Dataverse
- Project Online data exported to Files directory
- Project_{ProjectName}_published_mpp.xml files generated during Schedule-Migration copies to Files directory
- Lookup and choice data configured in Defaults.ps1
- Custom field mappings defined in ImportTasks.ps1
- This script and its 5.TaskField-Migration folder must remain relatively placed in the same parent directory as the 2.Project-and-Resource-Migration folder

Configuration:
- Edit Defaults.ps1 to configure default properties, lookup tables and choice fields
- Edit ImportTasks.ps1 to configure task field mappings
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$EnvironmentUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('E', 'W', 'Execute', 'WhatIf', IgnoreCase = $true)]
    [string]$ExecutionMode,
    
    [Parameter(Mandatory = $false)]
    [switch]$ForceLogin
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Check Prerequisites
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
    
    # Build parameter list to pass to pwsh
    $paramList = @()
    if ($EnvironmentUrl) { $paramList += "-EnvironmentUrl", $EnvironmentUrl }
    if ($ExecutionMode) { $paramList += "-ExecutionMode", $ExecutionMode }
    if ($ForceLogin) { $paramList += "-ForceLogin" }
    
    & pwsh.exe -File $MyInvocation.MyCommand.Path @paramList
    exit
} elseif ($psVersion.Major -lt 7 -or ($psVersion.Major -eq 7 -and $psVersion.Minor -lt 4)) {
    Write-Host "`nERROR: PowerShell 7.4 or later is required." -ForegroundColor Red
    Write-Host "Current version: $($psVersion.ToString())" -ForegroundColor Yellow
    Write-Host "Please download and install the latest PowerShell from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Yellow
    Write-Host "`nExiting..." -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ PowerShell version $($psVersion.ToString()) meets minimum requirement of 7.4" -ForegroundColor Green

# Check Azure CLI is installed
$azCli = Get-Command az -ErrorAction SilentlyContinue
if (-not $azCli) {
    Write-Host "`nERROR: Azure CLI is not installed." -ForegroundColor Red
    Write-Host "Please install the Azure CLI from: https://aka.ms/installazurecliwindows" -ForegroundColor Yellow
    Write-Host "`nExiting..." -ForegroundColor Red
    exit 1
}
$azCliVersion = (az version --query '"azure-cli"' -o tsv 2>$null)
Write-Host "  ✓ Azure CLI is installed (version $azCliVersion)" -ForegroundColor Green

Write-Host ""

. $PSScriptRoot\..\2.Project-and-Resource-Migration\Core.ps1
. $PSScriptRoot\..\2.Project-and-Resource-Migration\TableOperations.ps1
. $PSScriptRoot\..\2.Project-and-Resource-Migration\ExtendedTableOperations.ps1
. $PSScriptRoot\..\2.Project-and-Resource-Migration\CommonFunctions.ps1
. $PSScriptRoot\ImportTasks.ps1
. $PSScriptRoot\..\2.Project-and-Resource-Migration\LoadLookupAndChoiceData.ps1
# LoadLookupAndChoiceData.ps1 re-sources its own Defaults.ps1 from the 2. directory as a defensive
# measure. Re-source ours here to restore $LookupTablesToLoad and $ChoiceFieldsToLoad for this script.
. $PSScriptRoot\Defaults.ps1

# Start logging to file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logsFolder = Join-Path $PSScriptRoot "Logs"
# Create Logs folder if it doesn't exist
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}
$logFile = Join-Path $logsFolder "AltusTaskFieldMigration_$timestamp.txt"
Start-Transcript -Path $logFile -Append

Write-Host "Logging to: $logFile" -ForegroundColor Cyan
Write-Host ""

Write-Host "       " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host "`n" -BackgroundColor Black -NoNewline
Write-Host " Ʌ̲LTUS " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host "`n" -BackgroundColor Black -NoNewline
Write-Host "       " -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
Write-Host "`n" -BackgroundColor Black -NoNewline

Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
Write-Host "Task Field Migration Script" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

# Use parameter if provided, otherwise prompt user
if ([string]::IsNullOrWhiteSpace($EnvironmentUrl)) {
    $environmentUrl = Read-Host -Prompt "Enter the Dataverse environment URL (e.g., https://yourorg.crm.dynamics.com/)"
} else {
    $environmentUrl = $EnvironmentUrl
    Write-Host "`nEnvironment URL provided via parameter: $environmentUrl" -ForegroundColor Cyan
}

# Ensure the URL ends with a trailing slash
if (-not $environmentUrl.EndsWith('/')) {
    $environmentUrl += '/'
}

# Attempt connection
try {
    Connect $environmentUrl -Force:$ForceLogin
    
    # Verify connection succeeds (throws if not)
    Get-WhoAmI | Out-Null
    
    Write-Host "`nConnection to $environmentUrl successful!" -ForegroundColor Green
} catch {
    # Capture the error record immediately - $_ can be overwritten by cmdlets called inside the catch block.
    $connectError = $_

    # Detect the "user is not a member of the organization" error that occurs when cached
    # credentials belong to a different Dataverse environment. Force re-authentication and retry.
    $connectErrorMessage = $connectError.ToString()
    if ($connectErrorMessage -match 'not a member of the organization' -or
        $connectErrorMessage -match '80072560' -or
        $connectErrorMessage -match '401 \(Unauthorized\)' -or
        $connectErrorMessage -match 'Failed to connect to') {
        Write-Host "`nThe cached credentials do not have access to $environmentUrl." -ForegroundColor Yellow
        Write-Host "Please sign in with an account that has access to this environment.`n" -ForegroundColor Yellow

        try {
            Connect $environmentUrl -Force
            Get-WhoAmI | Out-Null
            Write-Host "`nConnection to $environmentUrl successful!" -ForegroundColor Green
        } catch {
            Write-Host "`nERROR: Failed to connect to $environmentUrl after re-authentication." -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            Write-Host "`nPlease verify:" -ForegroundColor Yellow
            Write-Host "  - The environment URL is correct" -ForegroundColor Yellow
            Write-Host "  - You have access to the environment" -ForegroundColor Yellow
            Write-Host "  - You signed in with the correct account" -ForegroundColor Yellow
            Write-Host "`nExiting..." -ForegroundColor Red
            Stop-Transcript
            exit 1
        }
    } else {
        Write-Host "`nERROR: Failed to connect to $environmentUrl" -ForegroundColor Red
        Write-Host "Error details: $connectError" -ForegroundColor Red
        Write-Host "`nPlease verify:" -ForegroundColor Yellow
        Write-Host "  - The environment URL is correct" -ForegroundColor Yellow
        Write-Host "  - You have access to the environment" -ForegroundColor Yellow
        Write-Host "  - You are authenticated with the correct account" -ForegroundColor Yellow
        Write-Host "`nExiting..." -ForegroundColor Red
        Stop-Transcript
        exit 1
    }
}

# Connection successful, display menu
Write-Host ""
Write-Host "IMPORTANT: Please ensure that the files to be imported are located in the 'Files' folder relative to this script." -ForegroundColor Yellow
Write-Host ""

# Use parameter if provided, otherwise prompt user
if ([string]::IsNullOrWhiteSpace($ExecutionMode)) {
    Write-Host "`nPlease select a mode of operation:" -BackgroundColor Gray -ForegroundColor Black
    Write-Host "`nE: Execute (will potentially create and/or update records in your Dataverse environment)" -BackgroundColor White -ForegroundColor Black
    Write-Host "W: What-If (simulates actions without actually making changes)" -BackgroundColor White -ForegroundColor Black
    Write-Host "`n"
    $executionMode = Read-Host -Prompt "Enter your choice (E or W)"
} else {
    $executionMode = $ExecutionMode
    Write-Host "`nExecution mode provided via parameter: $executionMode" -ForegroundColor Cyan
}

switch ($executionMode.ToUpper()) {
    { $_ -in "E", "EXECUTE" } {
        $executeScript = $true
        Write-Host "`n--------------------------------------------" -ForegroundColor Green
        Write-Host "Execution Mode: Execute" -ForegroundColor Green
        Write-Host "Script will make changes to your Dataverse environment." -ForegroundColor Green
        Write-Host "--------------------------------------------" -ForegroundColor Green
    }
    { $_ -in "W", "WHATIF" } {
        $executeScript = $false
        Write-Host "`n--------------------------------------------" -ForegroundColor Magenta
        Write-Host "Execution Mode: What-If" -ForegroundColor Magenta
        Write-Host "Script will simulate actions without making changes." -ForegroundColor Magenta
        Write-Host "--------------------------------------------" -ForegroundColor Magenta
    }
    default {
        Write-Host "`nInvalid selection. Please run the script again and choose a valid option." -ForegroundColor Red
        Stop-Transcript
        exit
    }
}

Write-Host "PRE LOADING LOOKUP AND CHOICE DATA..." -ForegroundColor Cyan
$global:lookupData = Get-LookupTableData
$global:choiceData = Get-ChoiceFieldData

Write-Host "`nStarting Task Field Migration..."
ImportTasks -ExecutionMode $executeScript

# Stop logging
Stop-Transcript
Write-Host "`nLog file saved to: $logFile" 