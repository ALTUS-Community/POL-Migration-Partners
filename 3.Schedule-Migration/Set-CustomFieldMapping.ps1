<#
.SYNOPSIS
    Wrapper for custom field mapping.
    Supports single field, bulk CSV, and batch processing of MPP files.
#>

[CmdletBinding()]
param(
    # Single field mode
    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [string]$customFieldName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [ValidatePattern("^(Text|Number|Date|Cost|Flag)\d+$")]
    [string]$localFieldTarget,

    # Bulk CSV mode
    [Parameter(Mandatory = $true, ParameterSetName = 'Batch')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Bulk')]
    [string]$MappingFile,

    # Batch mode
    [Parameter(Mandatory = $true, ParameterSetName = 'Batch')]
    [string]$MppFolder
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Force Windows PowerShell 5.1
if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "Restarting under Windows PowerShell 5.1..." -ForegroundColor Cyan
    
    $scriptPath = $MyInvocation.MyCommand.Path
    $argList = @("-Version", "5.1", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath)
    
    if ($MppFolder) {
        $argList += "-MppFolder", $MppFolder, "-MappingFile", $MappingFile
    } elseif ($MappingFile) {
        $argList += "-MappingFile", $MappingFile
    } else {
        $argList += "-customFieldName", $customFieldName, "-localFieldTarget", $localFieldTarget
    }
    
    & powershell.exe $argList
    exit
}

# We are now running under PowerShell 5.1
Write-Host "Running under Windows PowerShell 5.1" -ForegroundColor Green

# Load helpers
. (Join-Path $PSScriptRoot "CustomFieldMapping.ps1")
. (Join-Path $PSScriptRoot "Common.ps1")

# ------------------------------------------------------------------------------
# Batch mode
# ------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'Batch') {
    if (-not (Test-Path $MppFolder)) {
        throw "Folder not found: $MppFolder"
    }

    $mppFiles = Get-ChildItem -Path $MppFolder -Filter "*.mpp" -File |
    Where-Object { $_.Name -notlike "*_draft*" }

    if ($mppFiles.Count -eq 0) {
        Write-Warning "No .mpp files found in $MppFolder"
        exit
    }

    Write-Host "Found $($mppFiles.Count) MPP files. Starting batch mapping..." -ForegroundColor Cyan

    foreach ($mpp in $mppFiles) {
        Write-Host "`n=== Processing: $($mpp.Name) ===" -ForegroundColor Cyan

        try {
            Open-WinProj
            $global:WinProjApp.FileOpenEx($mpp.FullName, $false)
            Write-Host "Opened $($mpp.Name)" -ForegroundColor Green

            Set-CustomFieldMappings -MappingFile $MappingFile

            $global:WinProjApp.FileCloseEx('pjSave')
            Write-Host "Saved and closed $($mpp.Name)" -ForegroundColor Green
        } catch {
            Write-Error "Failed on $($mpp.Name): $($_.Exception.Message)"
        } finally {
            try { Close-AllOpenProjects } catch {}
        }
    }

    Write-Host "`nBatch mapping completed for $($mppFiles.Count) files." -ForegroundColor Green
    return
}

# ------------------------------------------------------------------------------
# Original modes
# ------------------------------------------------------------------------------
if ($PSCmdlet.ParameterSetName -eq 'Bulk') {
    Set-CustomFieldMappings -MappingFile $MappingFile
} else {
    Set-CustomFieldMapping -customFieldName $customFieldName -localFieldTarget $localFieldTarget
}