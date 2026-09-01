<#
.SYNOPSIS
  Run SharePoint export, import, or full migration for multiple site collections

.DESCRIPTION
  This script orchestrates the migration process for multiple SharePoint site collections.
  Supports three modes:
  - Export: Extract data from SharePoint to XML files
  - Import: Load XML files into Dynamics 365
  - Full: Export from SharePoint and immediately import to D365
  
  Each site collection gets its own output folder: ./Output/<SiteName>/<ProjectName>/
#>

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# ==============================================================================
# CONFIGURATION - CUSTOMIZE THESE SETTINGS
# ==============================================================================

# ------------------------------------------------------------------------------
# Site Collections to Migrate
# ------------------------------------------------------------------------------
# Define the SharePoint site collections to process.
# Each entry creates a separate output folder: Output/<FolderName>/<ProjectName>/
$SiteCollections = @(
  @{
    Url           = "https://contoso.sharepoint.com/sites/pwa"
    FolderName    = "PwaExample"      # Folder name for output
    ProjectFilter = @()     # Export all projects, or use patterns like @("*Example*", "Project*")
  }
)

# ------------------------------------------------------------------------------
# Dynamics 365 Settings (for Import Operations)
# ------------------------------------------------------------------------------
$D365Url = "https://orgname.crm.dynamics.com/"
$ImportParallelRequests = 4                          # Number of parallel import threads (1-10)
$ImportForce = $true                                 # $true = update existing records, $false = insert only

# ------------------------------------------------------------------------------
# POL Export Root (Step 1 output)
# ------------------------------------------------------------------------------
# Path to the POLExports directory produced in step 1. Use an existing absolute
# path if the exports are stored elsewhere. Leave empty to default to
# ./POLExports next to this script.
$PolExportRoot = "C:\Migration\Data"

# ------------------------------------------------------------------------------
# Authentication Settings
# ------------------------------------------------------------------------------
$ClientId = "YOUR-ENTRA-CLIENT-ID"

# ------------------------------------------------------------------------------
# File paths
# ------------------------------------------------------------------------------
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$MappingJsonPath = Join-Path $ScriptDir "export.config.json"
$CmtSchemaPath = Join-Path $ScriptDir "data_schema.xml"
$BaseOutputFolder = Join-Path $ScriptDir "Output"

# ==============================================================================
# END CONFIGURATION - Do not modify below this line unless you know what you're doing
# ==============================================================================

# Delegate execution to reusable runner so this file can be copied/repurposed
& (Join-Path $ScriptDir "Invoke-SharePointMigration.ps1") `
  -SiteCollections $SiteCollections `
  -D365Url $D365Url `
  -ImportParallelRequests $ImportParallelRequests `
  -ImportForce $ImportForce `
  -ClientId $ClientId `
  -MappingJsonPath $MappingJsonPath `
  -CmtSchemaPath $CmtSchemaPath `
  -BaseOutputFolder $BaseOutputFolder `
  -PolExportRoot $PolExportRoot `
  -ScriptDir $ScriptDir

exit $LASTEXITCODE
