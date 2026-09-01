<#
.SYNOPSIS
  Export SharePoint project-site lists to a CMT-compliant Data.xml using PnP.PowerShell, with correct lookup handling.

.DESCRIPTION
  Extracts project data from SharePoint project sites and converts to CMT-compliant XML format.
  
  IMPORTANT: Your user account must be a Site Collection Administrator on the target site collection.
  If you receive "Access is denied" errors, verify your Site Collection Admin role.

.PARAMETER SiteCollectionUrl
  The PWA site collection URL (e.g., https://tenant.sharepoint.com/sites/PWA)

.PARAMETER MappingJsonPath
  Path to JSON mapping file

.PARAMETER CmtSchemaPath
  Optional path to the CMT Schema.xml file for reference (Altus DataLoader no longer requires this)

.PARAMETER OutputFolder
  Folder to write Data.xml (and logs)

.PARAMETER ProjectFilter
  Optional wildcard pattern(s) to filter project sites by title. Supports multiple patterns.
  Examples: "Project A", "*2024*", "Project*", @("ProjectA", "ProjectB")

.PARAMETER ClientId
    Optional Entra ID Client ID for authentication. Configure it for the approved working tool.

.PARAMETER POLExportPath
  Optional path to Project Online export folder containing *_reporting.json files.
  If provided, project names and UIDs are sourced from the JSON export instead of SharePoint.
  Example: "C:\exports\POL\VNext"
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $SiteCollectionUrl,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string] $AzureEnvironment = 'Production',

    [Parameter(Mandatory = $false)]
    [string] $MappingJsonPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "export.config.json"),

    [Parameter(Mandatory = $false)]
    [string] $CmtSchemaPath = $null,

    [Parameter(Mandatory = $true)]
    [string] $OutputFolder,

    [Parameter(Mandatory = $false)]
    [string[]] $ProjectFilter = @(),

    [Parameter(Mandatory = $false)]
    [string] $ClientId = "YOUR-ENTRA-CLIENT-ID",

    [Parameter(Mandatory = $true)]
    [string] $POLExportPath
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Import PS5-safe relaunch helper and elevate to pwsh when required.
$relaunchHelperPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Compat\Relaunch.PS5Safe.ps1"
if (-not (Test-Path $relaunchHelperPath)) {
    Write-Error "Relaunch helper script not found: $relaunchHelperPath"
    $global:LASTEXITCODE = 1
    return
}
. $relaunchHelperPath
Invoke-RelaunchInPwshIfNeeded -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters

# Import helper functions
$commonHelpersPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "CommonFunctions.ps1"
if (-not (Test-Path $commonHelpersPath)) {
    Write-Error "Common helpers script not found: $commonHelpersPath"
    $global:LASTEXITCODE = 1
    return
}
. $commonHelpersPath

$script:hadError = $false

# ==============================================================================
# PATH RESOLUTION
# ==============================================================================

# Get the script's directory for resolving relative paths
$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path

# Resolve relative paths
$OutputFolder = Resolve-RelativePath -Path $OutputFolder -BasePath $scriptFolder
$POLExportPath = Resolve-RelativePath -Path $POLExportPath -BasePath $scriptFolder

# ==============================================================================
# INITIALIZATION
# ==============================================================================

Write-Host "`nExport Parameters:" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host "SiteCollectionUrl:        $SiteCollectionUrl" -ForegroundColor Gray
Write-Host "AzureEnvironment:         $AzureEnvironment" -ForegroundColor Gray
Write-Host "MappingJsonPath:          $MappingJsonPath" -ForegroundColor Gray
Write-Host "CmtSchemaPath:            $CmtSchemaPath" -ForegroundColor Gray
Write-Host "OutputFolder:             $OutputFolder" -ForegroundColor Gray
if ($ProjectFilter -and $ProjectFilter.Count -gt 0) {
    Write-Host "ProjectFilter:            $($ProjectFilter -join ', ')" -ForegroundColor Gray
}
Write-Host "ClientId:                 $ClientId" -ForegroundColor Gray
if ($POLExportPath) {
    Write-Host "POLExportPath:            $POLExportPath" -ForegroundColor Gray
}
Write-Host ""

# ==============================================================================
# PREREQUISITES
# ==============================================================================

Write-Host "Checking prerequisites..." -ForegroundColor Cyan
Test-PnPModule

if (-not (Test-Path $OutputFolder)) { 
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null 
}

# ==============================================================================
# LOAD CONFIGURATION
# ==============================================================================

Write-Host "Loading JSON mapping..." -ForegroundColor Cyan
if (-not (Test-Path $MappingJsonPath)) { 
    Write-Error "Mapping JSON not found."
    exit 1 
}

$Mapping = Get-Content -Raw -Path $MappingJsonPath | ConvertFrom-Json

# Load schema if provided (optional - for reference only)
$CmtSchema = $null
if ($CmtSchemaPath -and (Test-Path $CmtSchemaPath)) {
    Write-Host "Loading CMT schema for reference..." -ForegroundColor Gray
    [xml]$CmtSchema = Get-Content -Path $CmtSchemaPath
}

# Build project map from POL export if provided
$projectMap = @{}  # Map of project URL -> @{Name, UID}
if ($POLExportPath) {
    $projectMap = Build-ProjectMap -ExportPath $POLExportPath
}

# Initialize counters and logging
$script:totalProjectsProcessed = 0
$script:totalItemsExported = 0
$script:timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$script:logPath = Join-Path $OutputFolder "Export_$script:timestamp.log"
$script:logContent = @()

Write-LogMessage "Export started at $(Get-Date -Format o)"

# Create schema lookup only if schema was provided
$SchemaLookup = $null
if ($CmtSchema) {
    $SchemaLookup = New-SchemaLookup -Schema $CmtSchema
}

# ==============================================================================
# SHAREPOINT CONNECTIVITY
# ==============================================================================

Write-Host "Connecting to SharePoint (interactive as Site Collection Admin)..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteCollectionUrl -ClientId $ClientId -AzureEnvironment $AzureEnvironment -Interactive

# ==============================================================================
# WEB PROCESSING
# ==============================================================================

try {
    $webs = Get-FilteredProjectWebs -ProjectFilter $ProjectFilter -IncludeRootWeb $false -ProjectMap $projectMap

    foreach ($web in $webs) {
        # Use POL project name if available; otherwise use web title
        # Extract last segment of web URL for lookup
        $webUrlSegment = [uri]::UnescapeDataString(($web.Url -split '/')[-1])
        $projectName = $web.Title
        $projectId = $null
        if ($projectMap.ContainsKey($webUrlSegment)) {
            $projectName = $projectMap[$webUrlSegment].Name
            $projectId = $projectMap[$webUrlSegment].UID
        }

        Write-Host ""
        Write-Host "Processing web: $projectName ($($web.Url))" -ForegroundColor Cyan

        # ==============================================================================
        # GET PROJECT METADATA
        # ==============================================================================

        Connect-PnPOnline -Url $web.Url -ClientId $ClientId -AzureEnvironment $AzureEnvironment

        $projectItemCount = 0
        
        # ProjectId must come from POL map; we don't fall back to SharePoint
        if (-not $projectId) {            
            Write-LogWarning "Project '$projectName' ($($web.Url)) not found in POL export. Skipping."
            continue
        }
        
        Write-LogMessage "Using ProjectId from POL export: $projectId"        
        Write-LogMessage "Project: $projectName ($($web.Url))"

        # ==============================================================================
        # CREATE OUTPUT
        # ==============================================================================

        $dataXml = New-CmtDataXml
        $safeFolderName = $projectName -replace '[\\/:*?"<>|]', '_'
        $projectOutputFolder = Join-Path $OutputFolder $safeFolderName
        
        if (-not (Test-Path $projectOutputFolder)) { 
            New-Item -Path $projectOutputFolder -ItemType Directory | Out-Null 
        }

        # ==============================================================================
        # LIST PROCESSING
        # ==============================================================================
        
        # Extract scheme and domain from web URL (e.g., https://tenant.sharepoint.com)
        $uri = [System.Uri]$web.Url
        $rootUrl = "$($uri.Scheme)://$($uri.Authority)"

        foreach ($l in $Mapping.lists) {
            $spListTitle = $l.spListTitle
            $entityLogicalName = $l.entityLogicalName
            $projectLookupAttribute = $l.projectLookupAttribute
            $backlinkAttribute = $l.backlinkAttribute
            $columnMap = @($l.columnMap)

            # Get list
            try {
                $listUrl = "Lists/$spListTitle"
                $list = Get-PnPList -Identity $listUrl -Includes "DefaultDisplayFormUrl" -ErrorAction Stop 
                Write-LogMessage "  Found list: $spListTitle" -ForegroundColor Green
            } catch {
                Write-Warning "List '$spListTitle' not found. Skipping."
                Write-LogWarning "List '$spListTitle' not found in project '$projectName' ($($web.Url))."
                continue
            }

            # Get items
            try {
                $items = Get-PnPListItem -List $list -PageSize 5000 -ErrorAction Stop
                $itemCount = @($items).Count
                Write-LogMessage "  List '$spListTitle': $itemCount items" -ForegroundColor Cyan
                $projectItemCount += $itemCount
            } catch {
                Write-Warning "Error retrieving items from list '$spListTitle': $($_.Exception.Message)"
                Write-LogWarning "Error retrieving items from list '$spListTitle' in project '$projectName': $($_.Exception.Message)"
                $script:hadError = $true
                continue
            }

            # ==============================================================================
            # ITEM CONVERSION
            # ==============================================================================

            foreach ($item in $items) {
                # Build item URL using root domain + server-relative form URL + item ID
                # DefaultDisplayFormUrl is server-relative (e.g., "/sites/pwa/Example Project/Lists/Risks/DispForm.aspx")
                $itemUrl = "$($rootUrl)$($list.DefaultDisplayFormUrl)?ID=$($item.Id)"

                # Convert item to entity attributes
                $ctx = @{
                    columnMap              = $columnMap
                    backlinkAttribute      = $backlinkAttribute
                    projectLookupAttribute = $projectLookupAttribute
                    ProjectId              = $projectId
                    ItemUrl                = $itemUrl
                }
                
                $schemaFields = $null
                if ($SchemaLookup -and $SchemaLookup.ContainsKey($entityLogicalName)) {
                    $schemaFields = $SchemaLookup[$entityLogicalName]
                }

                $attrs = Convert-SpItemToEntity -Item $item.FieldValues -Ctx $ctx `
                    -ProjectGuid $projectId -ProjectName $projectName `
                    -EntityLogicalName $entityLogicalName -SchemaFieldLookup $schemaFields
            
                # Create deterministic GUID from SharePoint List ID and Item ID to ensure consistency across multiple exports
                # $deterministicId = New-HashGuid -InputString "$($list.Id)|$($item.Id)"
                # $attrs["_recordId"] = $deterministicId.ToString()
                $attrs["_recordId"] = $item.FieldValues["GUID"];
            
                Add-CmtEntityRecord -Doc $dataXml -EntityLogicalName $entityLogicalName -Attributes $attrs
            }
        }
    
        # ==============================================================================
        # SAVE PROJECT DATA
        # ==============================================================================

        $dataPath = Join-Path $projectOutputFolder "Data.xml"
        $dataXml.Save($dataPath)
        Write-Host "  Data.xml saved to: $dataPath" -ForegroundColor Green

        $script:totalProjectsProcessed++
        $script:totalItemsExported += $projectItemCount
    }

    # ==============================================================================
    # SUMMARY
    # ==============================================================================

    Write-Host ""
    Write-Host "`nExport complete!" -ForegroundColor Green
    Write-Host "Projects processed: $script:totalProjectsProcessed" -ForegroundColor Cyan
    Write-Host "Total items exported: $script:totalItemsExported" -ForegroundColor Cyan
} finally {
    # ==============================================================================
    # FINALIZATION
    # ==============================================================================

    $script:logContent += "Summary: Projects processed=$script:totalProjectsProcessed, Items exported=$script:totalItemsExported"
    $script:logContent += "Export finished at $(Get-Date -Format o)"
    Save-ExportLog
    
    $global:LASTEXITCODE = if ($script:hadError) { 1 } else { 0 }
}