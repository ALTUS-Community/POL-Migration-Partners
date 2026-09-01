<#
.SYNOPSIS
Exports all projects from PWA instance.

.DESCRIPTION
ExportAllProjects.ps1 exports all projects from a given PWA instance.
NOTE: This script borrows code from and depends upon the 'Project Online User Content Export and Delete script package' to run. 
See:
https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online

.PARAMETER Url
URL of the PWA instance.

.PARAMETER OutputDirectory
The location where user related data files will be stored.

.PARAMETER Region
The region of the tenant. Can be one of following values: "Default", "ITAR", "Germany" or "China"

.PARAMETER ProjectFilter
Optional wildcard pattern(s) to filter projects by name. Supports multiple patterns.
Examples: "Project A", "*2024*", "Project*", @("ProjectA", "ProjectB")

.PARAMETER ProjectOnlineOdataFilter
Optional OData filter string to apply when querying projects via the Project Online OData API.
Example: "ProjectStartDate gt datetime'2024-09-01T00:00:00Z'", "EnterpriseProjectTypeName eq 'Major Project'"

.EXAMPLE
.\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder
Exports all projects from the PWA site and saves it in the OutputFolder.

.EXAMPLE
    .\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder -ProjectFilter @("*Example*", "Project*")
Exports only projects whose names match one or more filter patterns.

.EXAMPLE
.\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder -ProjectOnlineOdataFilter "ProjectStartDate gt datetime'2024-09-01T00:00:00Z'"
Exports only projects that match the specified OData filter.

.NOTES
You need to be a site collection administrator and a PWA administrator to run this script.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Url,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Default", "ITAR", "Germany", "China")]
    [string] $Region = "Default",

    [Parameter(Mandatory = $false)]
    [string[]] $ProjectFilter = @(),

    [Parameter(Mandatory = $false)]
    [string] $ProjectOnlineOdataFilter = "",

    [Parameter(Mandatory = $false)]
    [switch] $OnPrem
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

. "$PSScriptRoot\Common.ps1"
. "$PSScriptRoot\CommonExtensions.ps1"
. "$PSScriptRoot\ExportDraftAndPublishedAsXML.ps1"


if (-not (ValidateDirectory $OutputDirectory)) {
    Write-Error "Error validating output location"
    return $false
}

Write-Host "Connecting..."
$proxy = GetPSIProxy -Url $Url -Region $Region -OnPrem $OnPrem
Write-Host "Proxy created"

if ($ProjectOnlineOdataFilter) {
    $ProjectOnlineOdataRequest = "Projects?`$filter=ProjectType ne 7 and $ProjectOnlineOdataFilter&`$select=ProjectId,ProjectName,ProjectOwnerId,ProjectOwnerName"
} else {
    $ProjectOnlineOdataRequest = "Projects?`$filter=ProjectType ne 7&`$select=ProjectId,ProjectName,ProjectOwnerId,ProjectOwnerName"
}

# Save on odata extract of all published projects (excluding the fake one - Timesheet Administrative Work Items)
SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request $ProjectOnlineOdataRequest -OutputDirectory $OutputDirectory -OutputfileName 'ReportingProjectList.json'

$projectsJson = (Get-Content -Path "$OutputDirectory\ReportingProjectList.json") | ConvertFrom-Json
$projectData = @($projectsJson.all.value)
$totalProjectCount = $projectData.Count

if ($ProjectFilter -and $ProjectFilter.Count -gt 0) {
    $projectData = @($projectData | Where-Object {
            $projectName = $_.ProjectName
            $isMatch = $false
            foreach ($filter in $ProjectFilter) {
                if ($projectName -like $filter) {
                    $isMatch = $true
                    break
                }
            }
            $isMatch
        })
}

$projectCount = $projectData.Count

if ($ProjectOnlineOdataFilter) {
    Write-Host "OData filter applied: $ProjectOnlineOdataFilter" -ForegroundColor Gray
}
if ($ProjectFilter -and $ProjectFilter.Count -gt 0) {
    Write-Host "Project filter applied: $($ProjectFilter -join ', ')" -ForegroundColor Gray
    Write-Host "Selected $projectCount of $totalProjectCount projects for export`n"
} else {
    Write-Host "Found $projectCount projects to export`n"
}

if ($projectCount -eq 0) {
    Write-Warning "No projects matched the supplied ProjectFilter. Nothing to export."
    return
}

$counter = 1
$percentComplete = 0

$siteId = (CallSPAPI -Proxy $proxy -BaseUrl $url -Request "Site?`$select=Id&`$top=1" -OnPrem $OnPrem) | Select-Object -ExpandProperty Id


Write-Host "Connecting MS Project to PWA..."
Connect-WinProjToProjectServer -pwaUrl $Url -siteId $siteId
Write-Host

# Track if we need to reset the connection
$script:needsConnectionReset = $false
$script:consecutiveFailures = 0

$projectData | ForEach-Object {
    $projName = $_.ProjectName
    $projUid = $_.ProjectId
    $ResourceUid = $_.ProjectOwnerId

    try {
        Write-Progress -Activity "Exporting Projects" -Status "Exporting project ($counter/$projectCount): `"$projName`"" -PercentComplete $percentComplete

        Write-Host "Processing Project: $projName ($projUid)"
        
        # Reset connection if needed from previous failure
        if ($script:needsConnectionReset) {
            Write-Warning "Resetting WinProj connection due to previous COM errors..."
            $resetSuccess = Reset-WinProjConnection -pwaUrl $Url -siteId $siteId
            if ($resetSuccess) {
                $script:needsConnectionReset = $false
                $script:consecutiveFailures = 0
                Write-Host "Connection reset successful, continuing exports..."
            } else {
                Write-Error "Failed to reset connection. Skipping MPP exports for: $projName"
                $script:consecutiveFailures++
            }
        }

        # Track if MPP exports fail for this project
        $mppExportFailed = $false
        
        ContinueOnError -ScriptBlock {
            try {
                Write-Host "Exporting draft schedule in MPP/XML formats..."
                Export-DraftProjectsAsXML -projectName $projName -projectGuid $projUid -exportFolder $OutputDirectory
            } catch {
                $mppExportFailed = $true
                if ($_.Exception.Message -match "unexpected error|RPC_E_|0x800706BA|0x80010001|COM") {
                    Write-Warning "COM error detected in draft export. Marking connection for reset."
                    $script:needsConnectionReset = $true
                    $script:consecutiveFailures++
                }
                throw
            }
        }

        ContinueOnError -ScriptBlock {
            try {
                Write-Host "Exporting published schedule in MPP/XML formats..."
                Export-PublishedProjectsAsXML -projectName $projName -projectGuid $projUid -exportFolder $OutputDirectory
            } catch {
                $mppExportFailed = $true
                if ($_.Exception.Message -match "unexpected error|RPC_E_|0x800706BA|0x80010001|COM") {
                    Write-Warning "COM error detected in published export. Marking connection for reset."
                    $script:needsConnectionReset = $true
                    $script:consecutiveFailures++
                }
                throw
            }
        }
        
        # If MPP export succeeded, reset failure counter
        if (-not $mppExportFailed) {
            $script:consecutiveFailures = 0
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting draft data for project: $projName..."
            $file = $OutputDirectory + "\Project_$($projName)_draft.json"
            $data = $proxy.UserDataExportDraftProject($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
            #Write-Host "Published data exported for project: $projName"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting published data for project: $projName..."
            $file = $OutputDirectory + "\Project_$($projName)_published.json"
            $data = $proxy.UserDataExportPubProject($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
            #Write-Host "Published data exported for project: $projName"
        }

        #Seems to consistently fail with DatabaseUndefinedError (?)
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting reporting data for project: $projName..."
            $file = $OutputDirectory + "\Project_$($projName)_reporting.json"
            $data = $proxy.UserDataExportReportingProject($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting tasks..."
            $file = $OutputDirectory + "\Project_$($projName)_reporting_Tasks.json"
            $data = $proxy.UserDataExportReportingProjectTasks($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting assignments..."
            $file = $OutputDirectory + "\Project_$($projName)_reporting_Assignments.json"
            $data = $proxy.UserDataExportReportingProjectAssignments($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting resources..."
            $file = $OutputDirectory + "\Project_$($projName)_reporting_Resources.json"
            $data = $proxy.UserDataExportReportingProjectResources($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting project baselines..."
            $file = $OutputDirectory + "\Project_$($projName)_reporting_Baselines.json"
            $data = $proxy.UserDataExportReportingProjectBaseline($ResourceUid, $projUid)
            $data = Repair-ProxyJson $data
            Set-Content -LiteralPath $file -Value $data -Encoding UTF8
            # OG code - fails when project name has special chars like '['
            # $data > $file
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting task timephased data..."
            $fileName = "Project_$($projName)_reporting_TaskTimephased.json"
            SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "TaskTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting assignment timephased data..." 
            $fileName = "Project_$($projName)_reporting_AssignmentTimephased.json"
            SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "AssignmentTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting task baseline timephased data..."
            $fileName = "Project_$($projName)_reporting_TaskBaselineTimephased.json"
            SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "TaskBaselineTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
        }

        ContinueOnError -ScriptBlock {
            Write-Host "...Exporting assignment baseline timephased data..."
            $fileName = "Project_$($projName)_reporting_AssignmentBaselineTimephased.json"
            SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "AssignmentBaselineTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
        }

    } catch {
        Write-Error "Error exporting project $projName. Error details: $_"
    } finally {
        Write-Host
        $counter++
        $percentComplete = ($counter / $projectCount) * 100
    }
}

Write-Host "Export complete!"

Write-Progress -Activity "Exporting Projects" -PercentComplete 100 -Completed