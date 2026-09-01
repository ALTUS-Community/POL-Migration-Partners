<#
.SYNOPSIS
Main orchestration script for migrating Project Online data to Altus (Dataverse).

.DESCRIPTION
This script coordinates the migration of projects and resources from Project Online 
to Altus Project Management. It handles:
- Connection to the Dataverse environment
- Loading lookup and choice field data
- Importing named resources (linked to system users)
- Importing generic resources
- Importing projects with custom field mappings

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

.PARAMETER Action
The import action to perform:
- 'ImportNamedResources': Import only named resources (linked to system users)
- 'ImportNamedAndGenericResources': Import both named and generic resources
- 'ImportProjects': Import projects with custom field mappings
If not provided, you will be prompted to select an action interactively.

.PARAMETER ForceLogin
When specified, forces a fresh Azure CLI login prompt even if valid cached credentials exist.
Use this when you need to authenticate as a different user with higher permissions in the
target environment (e.g. an Altus Admin rather than a read-only account).

.EXAMPLE
.\AltusPOLMigration.ps1

Runs interactively, prompting for all parameters.

.EXAMPLE
.\AltusPOLMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode WhatIf -Action ImportProjects

Runs in test mode to import projects without making actual changes.

.EXAMPLE
.\AltusPOLMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode Execute -Action ImportNamedResources

Imports named resources and creates records in Dataverse.

.EXAMPLE
.\AltusPOLMigration.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com/" -ExecutionMode Execute -Action ImportProjects -ForceLogin

Forces a fresh login prompt before importing projects, allowing you to authenticate as a
different user (e.g. one with Altus Admin permissions).

.NOTES
Prerequisites:
- PowerShell 7.4 or later
- Azure CLI (https://aka.ms/installazurecliwindows) — used for authentication to Dataverse
- Project Online data exported to Files directory
- Lookup and choice data configured in Defaults.ps1
- Custom field mappings defined in ImportProjects.ps1 and ImportResources.ps1

Configuration:
- Edit Defaults.ps1 to configure default properties, lookup tables and choice fields
- Edit ImportProjects.ps1 to configure project field mappings
- Edit ImportResources.ps1 to configure resource field mappings
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$EnvironmentUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('E', 'W', 'Execute', 'WhatIf', IgnoreCase = $true)]
    [string]$ExecutionMode,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('ImportNamedResources', 'ImportNamedAndGenericResources', 'ImportProjects', IgnoreCase = $true)]
    [string]$Action,

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
    if ($Action) { $paramList += "-Action", $Action }
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

. $PSScriptRoot\Core.ps1
. $PSScriptRoot\TableOperations.ps1
. $PSScriptRoot\CommonFunctions.ps1
. $PSScriptRoot\ImportResources.ps1
. $PSScriptRoot\ImportProjects.ps1
. $PSScriptRoot\LoadLookupAndChoiceData.ps1

# Start logging to file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logsFolder = Join-Path $PSScriptRoot "Logs"
# Create Logs folder if it doesn't exist
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}
$logFile = Join-Path $logsFolder "AltusPOLMigration_$timestamp.txt"
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
Write-Host "Project Online Migration Script" -ForegroundColor Cyan
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

# Use parameter if provided, otherwise prompt user
if ([string]::IsNullOrWhiteSpace($Action)) {
    Write-Host "`nPlease select an action:" -BackgroundColor Gray -ForegroundColor Black
    Write-Host "`n1: Import Resources (Named only)" -BackgroundColor White -ForegroundColor Black
    Write-Host "2: Import Resources (Named and Generic)" -BackgroundColor White -ForegroundColor Black
    Write-Host "3: Import Projects" -BackgroundColor White -ForegroundColor Black
    Write-Host "Q: Quit" -BackgroundColor White -ForegroundColor Black
    Write-Host "`n"
    $actionChoice = Read-Host -Prompt "Enter your choice (1, 2, 3 or Q)"
} else {
    $actionChoice = $Action
    Write-Host "`nAction provided via parameter: $actionChoice" -ForegroundColor Cyan
}

Write-Host "PRE LOADING LOOKUP AND CHOICE DATA..." -ForegroundColor Cyan
$global:lookupData = Get-LookupTableData
$global:choiceData = Get-ChoiceFieldData

switch ($actionChoice.ToUpper()) {
    { $_ -in "1", "IMPORTNAMEDRESOURCES" } {
        Write-Host "`nStarting Import Resources (Named only)..." 
        ImportResources -Mode 'NamedOnly' -ExecutionMode $executeScript
    }
    { $_ -in "2", "IMPORTNAMEDANDGENERICRESOURCES" } {
        Write-Host "`nStarting Import Resources (Named and Generic)..."
        ImportResources -Mode 'NamedAndGeneric' -ExecutionMode $executeScript
    }
    { $_ -in "3", "IMPORTPROJECTS" } {
        Write-Host "`nStarting Import Projects..." 
        ImportProjects -ExecutionMode $executeScript
    }
    "Q" {
        Write-Host "`nExiting..." -ForegroundColor Yellow
        Stop-Transcript
        exit
    }
    default {
        Write-Host "`nInvalid selection. Please run the script again and choose a valid option." -ForegroundColor Red
        Stop-Transcript
        exit
    }
}

# Stop logging
Stop-Transcript
Write-Host "`nLog file saved to: $logFile" 