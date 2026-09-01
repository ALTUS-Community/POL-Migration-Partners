<#
.SYNOPSIS
  Run SharePoint export/import/document operations for a set of site collections.

.DESCRIPTION
  This script contains the core interactive prompt and per-site processing moved
  out of -Run-Migration.ps1 so it can be reused. Callers should construct the
  configuration (site collection list and other settings) and pass them in.
#>

param(
    [Parameter(Mandatory = $true)]
    [array]$SiteCollections,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string] $AzureEnvironment = 'Production',

    [Parameter(Mandatory = $false)]
    [string]$D365Url = "",

    [Parameter(Mandatory = $false)]
    [int]$ImportParallelRequests = 4,

    [Parameter(Mandatory = $false)]
    [bool]$ImportForce = $true,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "YOUR-ENTRA-CLIENT-ID",

    [Parameter(Mandatory = $false)]
    [string]$MappingJsonPath = "",

    [Parameter(Mandatory = $false)]
    [string]$CmtSchemaPath = "",

    [Parameter(Mandatory = $false)]
    [string]$BaseOutputFolder = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Output"),

    [Parameter(Mandatory = $false)]
    [string]$PolExportRoot = "",

    [Parameter(Mandatory = $false)]
    [string]$ScriptDir = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Resolve POL export root if not provided
$PolExportRootSpecified = $false
if ([string]::IsNullOrWhiteSpace($PolExportRoot)) {
    $PolExportRoot = Join-Path $ScriptDir "POLExports"
} else {
    $PolExportRootSpecified = $true
}

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "SharePoint to Altus Migration Tool" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

# Display configured site collections
Write-Host "Configured Site Collections:" -ForegroundColor Yellow
foreach ($site in $SiteCollections) {
    Write-Host "  - $($site.FolderName): " -NoNewline -ForegroundColor White
    Write-Host "$($site.Url)" -ForegroundColor Gray
    if ($site.ProjectFilter -and $site.ProjectFilter.Count -gt 0) {
        Write-Host "    Filter: $($site.ProjectFilter -join ', ')" -ForegroundColor DarkGray
    }
}
Write-Host ""

# Display Dynamics 365 target
Write-Host "Dynamics 365 Target:" -ForegroundColor Yellow
Write-Host "  - Environment: " -NoNewline -ForegroundColor White
Write-Host "$D365Url" -ForegroundColor Gray
Write-Host "  - Parallel Requests: $ImportParallelRequests" -ForegroundColor Gray
Write-Host "  - Force Update: $ImportForce" -ForegroundColor Gray
Write-Host ""

Write-Host "Select operation mode:" -ForegroundColor Yellow
Write-Host "  [1] Export Lists - Extract data from SharePoint to XML files" -ForegroundColor White
Write-Host "  [2] Import Lists - Load existing XML files into Dynamics 365" -ForegroundColor White
Write-Host "  [3] Export Documents - Extract document attachments from SharePoint lists" -ForegroundColor White
Write-Host "  [Q] Quit" -ForegroundColor Gray
Write-Host ""

do {
    $choice = Read-Host "Enter your choice (1, 2, 3, or Q)"
    $validChoice = $choice -match '^[123Qq]$'
    if (-not $validChoice) {
        Write-Host "Invalid choice. Please enter 1, 2, 3, or Q" -ForegroundColor Red
    }
} while (-not $validChoice)

if ($choice -match '^[Qq]$') {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit 0
}

$operationMode = switch ($choice) {
    "1" { "ExportData" }
    "2" { "ImportData" }
    "3" { "ExportDocuments" }
}

Write-Host "`nSelected mode: $operationMode" -ForegroundColor Green
Write-Host ""

# Prompt for D365 url if needed
if ([string]::IsNullOrWhiteSpace($D365Url) -and $operationMode -eq "ImportData") {
    Write-Host "No Dynamics 365 URL configured." -ForegroundColor Yellow
    Write-Host ""
    $D365Url = Read-Host "Enter the Dynamics 365 environment URL (e.g., https://orgname.crm.dynamics.com)"
    if ([string]::IsNullOrWhiteSpace($D365Url)) {
        Write-Host "No D365 URL provided. Operation cancelled." -ForegroundColor Red
        exit 0
    }
    Write-Host ""
    Write-Host "Using D365 environment: $D365Url" -ForegroundColor Green
    Write-Host ""
}

# If no site collections configured, prompt for a single one when exporting
if ($SiteCollections.Count -eq 0 -and ($operationMode -eq "ExportData" -or $operationMode -eq "ExportDocuments")) {
    Write-Host "No site collections configured." -ForegroundColor Yellow
    Write-Host ""
    $siteUrl = Read-Host "Enter the source SharePoint site collection URL (e.g., https://tenant.sharepoint.com/sites/pwa)"
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Write-Host "No URL provided. Operation cancelled." -ForegroundColor Red
        exit 0
    }
    Write-Host ""
    $folderName = Read-Host "Enter the folder name for POLExports lookup (e.g., PWA_Main, vNext)"
    if ([string]::IsNullOrWhiteSpace($folderName)) {
        Write-Host "No folder name provided. Operation cancelled." -ForegroundColor Red
        exit 0
    }
    $SiteCollections = @(
        @{
            Url           = $siteUrl
            FolderName    = $folderName
            ProjectFilter = @()
        }
    )
    Write-Host ""
    Write-Host "Using site collection: $siteUrl" -ForegroundColor Green
    Write-Host "Output folder: $folderName" -ForegroundColor Gray
    Write-Host "POL Export path: .\POLExports\$folderName" -ForegroundColor Gray
    Write-Host ""
}

# Process each site collection
$totalSites = $SiteCollections.Count
$processedSites = 0
$results = @()

Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "$operationMode - Multi-Site Collection Migration" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Processing $totalSites site collection(s)`n" -ForegroundColor White

foreach ($site in $SiteCollections) {
    $processedSites++
    $siteUrl = $site.Url
    $folderName = $site.FolderName
    $projectFilter = $site.ProjectFilter

    if ($PolExportRootSpecified) {
        $polExportPath = $PolExportRoot
        Write-Host "Using provided POL export root: $PolExportRoot" -ForegroundColor Gray
    } else {
        $polExportPath = Join-Path $PolExportRoot $folderName
        Write-Host "Using default POL export root: $PolExportRoot" -ForegroundColor Gray
    }

    Write-Host "  POL Export Path: $polExportPath" -ForegroundColor Gray
    Write-Host "`n[$processedSites/$totalSites] Processing: $folderName" -ForegroundColor Yellow
    Write-Host "  URL: $siteUrl" -ForegroundColor Gray

    # Create site-specific output folder (only for list export/import)
    $siteOutputFolder = Join-Path $BaseOutputFolder $folderName
    if ($operationMode -ne "ExportDocuments") {
        if (-not (Test-Path $siteOutputFolder)) {
            New-Item -Path $siteOutputFolder -ItemType Directory -Force | Out-Null
            Write-Host "  Created output folder: $siteOutputFolder" -ForegroundColor Gray
        }
    }

    # Create document-specific output folder only when exporting documents
    if ($operationMode -eq "ExportDocuments") {
        $documentOutputFolder = Join-Path (Join-Path $BaseOutputFolder "Documents") $folderName
        if (-not (Test-Path $documentOutputFolder)) {
            New-Item -Path $documentOutputFolder -ItemType Directory -Force | Out-Null
            Write-Host "  Created document output folder: $documentOutputFolder" -ForegroundColor Gray
        }
    }

    $siteSuccess = $true
    $siteStartTime = Get-Date
    
    if ($operationMode -eq "ExportData" -or $operationMode -eq "ExportDocuments") {
        if ($projectFilter -and $projectFilter.Count -gt 0) {
            Write-Host "  Project Filter: $($projectFilter -join ', ')" -ForegroundColor Gray
        }

        try {
            if ($operationMode -eq "ExportData") {
                $exportParams = @{
                    SiteCollectionUrl = $siteUrl
                    AzureEnvironment  = $AzureEnvironment
                    MappingJsonPath   = $MappingJsonPath
                    OutputFolder      = $siteOutputFolder
                    ClientId          = $ClientId
                    POLExportPath     = $polExportPath
                }
                if ($CmtSchemaPath -and (Test-Path $CmtSchemaPath)) { $exportParams.CmtSchemaPath = $CmtSchemaPath }
                if ($projectFilter -and $projectFilter.Count -gt 0) { $exportParams.ProjectFilter = $projectFilter }

                Write-Host "  Starting list export..." -ForegroundColor Cyan
                $exportStartTime = Get-Date
                & (Join-Path $ScriptDir "1-export-lists.ps1") @exportParams
            } elseif ($operationMode -eq "ExportDocuments") {
                $exportParams = @{
                    SiteCollectionUrl = $siteUrl
                    AzureEnvironment  = $AzureEnvironment
                    OutputFolder      = $documentOutputFolder
                    ClientId          = $ClientId
                    POLExportPath     = $polExportPath
                }
                if ($projectFilter -and $projectFilter.Count -gt 0) { $exportParams.ProjectFilter = $projectFilter }

                Write-Host "  Starting document export..." -ForegroundColor Cyan
                $exportStartTime = Get-Date
                & (Join-Path $ScriptDir "3-export-documents.ps1") @exportParams
            }

            $exportDuration = (Get-Date) - $exportStartTime
            if ($LASTEXITCODE -ne 0) {
                $siteSuccess = $false
                Write-Host "  Export failed!" -ForegroundColor Red
            } else {
                Write-Host "  Export completed in $($exportDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Green
            }
        } catch {
            Write-Host "  ERROR during export: $($_.Exception.Message)" -ForegroundColor Red
            $siteSuccess = $false
        }
    }

    if ($operationMode -eq "ImportData" -and $siteSuccess) {
        $importParams = @{
            D365Url          = $D365Url
            DataFolder       = $siteOutputFolder
            ParallelRequests = $ImportParallelRequests
            Force            = $ImportForce
            POLExportPath    = $polExportPath
        }
        if ($projectFilter -and $projectFilter.Count -gt 0) { $importParams.ProjectFilter = $projectFilter }

        try {
            Write-Host "  Starting import to D365..." -ForegroundColor Cyan
            $importStartTime = Get-Date
            & (Join-Path $ScriptDir "2-import-lists.ps1") @importParams
            $importDuration = (Get-Date) - $importStartTime
            if ($LASTEXITCODE -ne 0) {
                $siteSuccess = $false
                Write-Host "  Import failed!" -ForegroundColor Red
            } else {
                Write-Host "  Import completed in $($importDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Green
            }
        } catch {
            Write-Host "  ERROR during import: $($_.Exception.Message)" -ForegroundColor Red
            $siteSuccess = $false
        }
    } elseif ($operationMode -eq "ImportData" -and -not $siteSuccess) {
        Write-Host "  Skipping import due to export failure" -ForegroundColor Yellow
    }

    $siteDuration = (Get-Date) - $siteStartTime
    $status = if ($siteSuccess) { "Success" } else { "Failed" }

    $results += [PSCustomObject]@{
        ExportFolder = $folderName
        Url          = $siteUrl
        Operation    = $operationMode
        Status       = $status
        Duration     = $siteDuration.ToString("hh\:mm\:ss")
        OutputFolder = if ($operationMode -eq "ExportDocuments") { $documentOutputFolder } else { $siteOutputFolder }
    }

    Write-Host "  Total time: $($siteDuration.ToString('hh\:mm\:ss'))" -ForegroundColor $(if ($siteSuccess) { "Green" } else { "Red" })
}

# Display summary
Write-Host "" 
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host "Migration Summary" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Cyan
Write-Host ""

$results | Format-Table -AutoSize

$successCount = ($results | Where-Object { $_.Status -eq "Success" }).Count
$failedCount = ($results | Where-Object { $_.Status -ne "Success" }).Count

Write-Host "Total Sites:     $totalSites" -ForegroundColor White
Write-Host "Successful:      $successCount" -ForegroundColor Green
Write-Host "Failed:          $failedCount" -ForegroundColor $(if ($failedCount -gt 0) { "Red" } else { "Gray" })
Write-Host ""

# Export summary to JSON
$summaryPath = Join-Path $BaseOutputFolder "migration-summary.json"
$results | ConvertTo-Json -Depth 5 | Out-File -FilePath $summaryPath -Encoding UTF8
Write-Host "Summary exported to: $summaryPath" -ForegroundColor Gray

return
