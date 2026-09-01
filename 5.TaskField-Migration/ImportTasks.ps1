# ============================================================================
# TASK FIELD IMPORT SCRIPT
# ============================================================================
# 
# USER CONFIGURATION: See "USER CONFIGURATION (EDIT HERE)" section below
#
# ============================================================================

. $PSScriptRoot\Defaults.ps1
. $PSScriptRoot\TaskMappingHelpers.ps1
# ==============================================================================
# LOGGING HELPERS (timestamps match stage 4 SharePoint scripts)
# ==============================================================================

function Write-Log {
    param([string]$Message, [string]$ForegroundColor = 'Gray')
    Write-Host "[$( Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $ForegroundColor
}

function Write-LogWarning {
    param([string]$Message)
    Write-Host "[$( Get-Date -Format 'HH:mm:ss')] [WARNING] $Message" -ForegroundColor Yellow
}

function Write-LogError {
    param([string]$Message)
    Write-Host "[$( Get-Date -Format 'HH:mm:ss')] [ERROR] $Message" -ForegroundColor Red
}


# ------------------------------
# USER CONFIGURATION (EDIT HERE)
# ------------------------------

<#
.SYNOPSIS
Maps Project Online task data to Dataverse task fields.

.DESCRIPTION
This function is the primary customisation point for mapping Project Online task fields 
to your Dataverse task table columns. Add your custom field mappings here to control 
how data flows from Project Online to Dataverse during the import process.

The function is called for each task being imported.

.PARAMETER Task
The hashtable representing the Dataverse task record being updated. 
Add field mappings to this hashtable using Dataverse logical column names.

.PARAMETER TaskName
The name of the task.

.PARAMETER TaskGuid
The unique GUID of the task from Project Online.

.PARAMETER ReportingTask
The Project Online reporting data for this task (if available). Contains OOTB fields 
and custom fields exported from Project Online. 

.EXAMPLE
# Map a text field from Project Online
$textValue = Get-ReportingCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName 'Division'
if ($textValue) { $Task['cr62c_division'] = $textValue }

.EXAMPLE
# Map a lookup field
Set-TaskLookupField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Department' -LookupName 'Department' -DataverseFieldName 'cr62c_Department'

.EXAMPLE
# Map a choice field
Set-TaskChoiceField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Status' -ChoiceName 'TaskStatus' -DataverseFieldName 'cr62c_status'

.NOTES
- Use Dataverse logical column names (lowercase) for field mappings
- All target columns must already exist in Dataverse
- Lookup and choice data must be pre-loaded via Defaults.ps1
- See the commented examples within the function for more mapping patterns
#>
function Add-TaskDataverseFieldMappings {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Task,

        [Parameter(Mandatory = $true)]
        [string]$TaskName,

        [Parameter(Mandatory = $true)]
        [string]$TaskGuid,

        [Parameter(Mandatory = $false)]
        $ReportingTask
    )

    # Add Task -> Dataverse column mappings here.
    # Use Dataverse logical column names (lowercase). Columns must already exist.

    # --- OOTB fields (from $ReportingTask or $TaskName) ---
    # $Task['cr_task_text_ootb'] = $TaskName
    # $Task['cr_task_datetime_ootb'] = ($ReportingTask.TaskEarlyFinish -as [datetime])
    # $Task['cr_task_dateonly_ootb'] = ($ReportingTask.TaskEarlyFinish -as [datetime]).ToString("yyyy-MM-dd")
    # $Task['cr_task_whole_ootb'] = ($ReportingTask.TaskPercentWorkCompleted -as [int])
    # $Task['cr_task_decimal_ootb'] = ($ReportingTask.TaskTotalSlack -as [decimal])

    # --- Enterprise Custom Fields (from $ReportingTask.CustomFields[]) ---
    # $textValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName 'Your Text Field'
    # if ($textValue) { $Task['cr_task_text_custom'] = [string]$textValue }
    #
    # $dateValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName 'Your Date Field'
    # if ($dateValue) { $Task['cr_task_datetime_custom'] = ($dateValue -as [datetime]) }
    # if ($dateValue) { $Task['cr_task_dateonly_custom'] = ($dateValue -as [datetime]).ToString("yyyy-MM-dd") }

    # --- Lookup Field Example ---
    # Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file 
    #       so that the correct corresponding lookup value can be populated.
    # Also note that -DataverseFieldName is CASE SENSITIVE when identifying the Dataverse Schema Name for the lookup field
    #Set-TaskLookupField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Task Schedule KPI' -LookupName 'TaskScheduleKPI' -DataverseFieldName 'cr9f9_TaskScheduleKPI'

    #--- Choice Field Example ---
    # Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding choice value can be populated.
    # If mapping to a multi-select choice field, set -IsMulti to $true
    # Set-TaskChoiceField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Investment Category' -ChoiceName 'Investment Category' -DataverseFieldName 'sensei_investmentcategory'

    #--- Many-To-Many Relationship Example ---
    # Note: To map to a Dataverse many-to-many relationship, you must ensure the related table data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding related records can be populated.
    # This example demonstrates handling multi-select Project Online lookup fields that populate a Dataverse relationship table.
    # The relationship table is populated AFTER the project is created/updated using the Add-ToCollection function.
    # Note that the relationship table is populated additively only, so existing records in the relationship will be retained.
    #Set-TaskManyToManyRelationship -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Task Work KPI' -LookupName 'TaskWorkKPI' -RelationshipName 'cr9f9_TaskWorkKPI_sensei_task_sensei_task' -RelatedTableCollectionName 'cr9f9_taskworkkpis' -RelatedTablePrimaryKey 'cr9f9_taskworkkpiid'
}

# <#
# .SYNOPSIS
# Creates the Dataverse project record body with default fields and custom mappings.

# .DESCRIPTION
# Helper function that constructs the complete hashtable for creating or updating a task 
# in Dataverse. It initializes required OOTB fields (name, task type, external task ID) 
# and then calls Add-TaskDataverseFieldMappings to add any custom field mappings defined 
# by the user.

# This function is called internally by the import process and should not typically need 
# to be modified.

# .PARAMETER TaskGuid
# The unique GUID of the task from Project Online.

# .PARAMETER ReportingTask
# The Project Online reporting data for this task.

# .OUTPUTS
# Returns a hashtable containing all field mappings ready to be used in a Dataverse 
# New-Record or Update-Record operation.

# .EXAMPLE
# $taskBody = Get-TaskDataverseBody -TaskGuid "12345678-1234-1234-1234-123456789012" -ReportingTask $reportingData

# .NOTES
# - This function sets the required OOTB Altus fields
# - Custom field mappings are added via Add-TaskDataverseFieldMappings
# #>
function Get-TaskDataverseBody {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskGuid,

        [Parameter(Mandatory = $false)]
        $ReportingTask
    )

    $taskName = if ($ReportingTask -and $ReportingTask.PSObject.Properties.Name -contains 'TaskName') { $ReportingTask.TaskName } else { '' }

    $task = @{}

    Add-TaskDataverseFieldMappings -Task $task -TaskName $taskName -TaskGuid $TaskGuid -ReportingTask $ReportingTask
    return $task
}

# ====== Don't edit below this line ======

<#
.SYNOPSIS
Imports tasks into Dataverse.

.DESCRIPTION
The ImportTasks function handles the import of tasks from JSON files into Dataverse.

.PARAMETER ExecutionMode
Controls whether to actually execute the import or run in test mode. When $true, records are created. When $false, runs in test mode. Default is $false.

.EXAMPLE
ImportTasks -ExecutionMode $true
This example imports tasks into Dataverse, updating records.
#>

function ImportTasks {
    param(
        [Parameter(Mandatory = $false)]
        [bool]$ExecutionMode = $false
    )

    $executionModeText = if ($ExecutionMode) { "Execute" } else { "What-If" }
    Write-Host "Import Task Fields (ExecutionMode: $executionModeText)" -ForegroundColor Green    
    
    Invoke-DataverseCommands {
        $nProjectsIterated = 0
        $nProjectsWithIncompleteFiles = 0
        $nTasksUpdated = 0
        $nTasksNotRequiringUpdate = 0
        $nErrored = 0
        $projectSummaries = @()

        # Read all Projects from Altus with a sensei_externalprojectid set
        $projectsWithExternalId = Get-ProjectsWithExternalId
        Write-Host "Retrieved $($projectsWithExternalId.Count) Projects with External IDs from Dataverse."

        $projectDesktopProjects = $projectsWithExternalId | Where-Object {
            $null -ne $_.sensei_externalproject_project_sensei_pro -and 
            $_.sensei_externalproject_project_sensei_pro.Count -gt 0 -and 
            $_.sensei_externalproject_project_sensei_pro._sensei_externalsystem_value -eq $ProjectDesktopExternalSystemId
        }

        $projectsWithExternalIdButNoExternalProjectReference = $projectsWithExternalId | Where-Object {
            $null -eq $_.sensei_externalproject_project_sensei_pro -or
            $_.sensei_externalproject_project_sensei_pro.Count -eq 0
        }

        Write-Host "Retrieved $($projectDesktopProjects.Count) Project Desktop Projects from Dataverse."
        Write-Host "Retrieved $($projectsWithExternalIdButNoExternalProjectReference.Count) Projects with External IDs but no External Project reference from Dataverse."
        
        #Read Solution version from Altus
        $solutionVersion = Get-SolutionVersion
        if ($solutionVersion) {
            Write-Host "Altus Solution Version: $solutionVersion"
        }
        else {
            Write-Host "Altus Solution: Not found" -ForegroundColor Red
            return
        }

        #Read Organization info from environment
        $orgName = Get-OrgName
        if ($orgName) {
            Write-Host "Organization Name: $orgName"
        }
        else {
            Write-Host "Organization Name: Not found" -ForegroundColor Red
            return
        }

        #Get EnvironmentId for this environment
        $envId = Get-EnvironmentId
        if ($envId) {
            Write-Host "Environment ID: $envId"
        }
        else {
            Write-Host "Environment ID: Not found" -ForegroundColor Red
            return
        }

        $filesPath = Join-Path $PSScriptRoot "Files"
        if (-not (Test-Path $filesPath)) {
            New-Item -Path $filesPath -ItemType Directory -Force | Out-Null
        }
        $mppsInFolder = Get-ChildItem -Path $filesPath -Filter "Project_*_published.mpp"

        if ($mppsInFolder.Count -eq 0) {
            Write-Host "No MPP files found in 'Files' folder. Please ensure the files are present." -ForegroundColor Red
            return
        }

        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Processing Projects..." -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan

        foreach ($mpp in $mppsInFolder) {
            Write-Host "`n-- Processing Project with MPP file: $($mpp.Name)..." -ForegroundColor Cyan
            
            try {
                # Read the corresponding JSON file with the same name
                $jsonFileName = [System.IO.Path]::GetFileNameWithoutExtension($mpp.Name) + ".json"
                $jsonFilePath = Join-Path $filesPath $jsonFileName
                
                if (Test-Path -LiteralPath $jsonFilePath) {
                    Write-Host "Reading JSON file: $jsonFileName" -ForegroundColor Gray
                    $projectData = Get-Content -LiteralPath $jsonFilePath -Raw | ConvertFrom-Json
                    $projectGuid = $projectData.NewDataSet.Project.ProjectUId
                    $projectName = $projectData.NewDataSet.Project.ProjectName
                    Write-Host "Project: $projectName ($projectGuid)" -ForegroundColor Gray

                    # Check if Project exists in Dataverse (based on sensei_externalprojectid match)
                    $existingProject = $projectDesktopProjects | Where-Object { $_.sensei_externalprojectid -eq "ProjectDesktop_$projectGuid" }
                    #check if there is a sensei_project record with the sensei_externalprojectid but no external project reference
                    $orphanedProject = $projectsWithExternalIdButNoExternalProjectReference | Where-Object { $_.sensei_externalprojectid -eq "ProjectDesktop_$projectGuid" }
                    if ($orphanedProject) {
                        Write-Host "Found orphaned Project record for $projectName ($projectGuid) in Dataverse. Ensure you first run the Import Projects script." -ForegroundColor Yellow
                    }

                    # If exists, proceed with task updates
                    if ($existingProject) {
                        $nProjectsIterated++
                        $nProjectTasksUpdated = 0
                        $nProjectTasksSkipped = 0
                        $nProjectTasksErrored = 0

                        # --- Load Task Metadata (Required for Custom Field Mapping) ---
                        # The _reporting_Tasks.json contains task-level reporting fields (including custom fields) useful for mapping into Dataverse.
                        $reportingTasks = $null
                        $reportingTasksJsonFileName = $mpp.Name.Replace("_published.mpp", "_reporting_Tasks.json")
                        $reportingTasksJsonPath = Join-Path $filesPath $reportingTasksJsonFileName

                        # The _published.xml file contains the mapping of Project Online Task GUIDs to integer UIDs, which is required to link the reporting data to the correct tasks and to find the Altus Task ID for updating the correct Task records in Dataverse.
                        $publishedXml = $null
                        $publishedXmlFileName = $mpp.Name.Replace(".mpp", ".xml")
                        $publishedXmlPath = Join-Path $filesPath $publishedXmlFileName

                        $publishedMppXml = $null
                        $publishedMppXmlFileName = $mpp.Name.Replace(".mpp", "_mpp.xml")
                        $publishedMppXmlPath = Join-Path $filesPath $publishedMppXmlFileName

                        $missingFiles = @()
                        if (-not (Test-Path -LiteralPath $reportingTasksJsonPath)) { $missingFiles += $reportingTasksJsonFileName }
                        if (-not (Test-Path -LiteralPath $publishedXmlPath)) { $missingFiles += $publishedXmlFileName }
                        if (-not (Test-Path -LiteralPath $publishedMppXmlPath)) { $missingFiles += $publishedMppXmlFileName }

                        if ($missingFiles.Count -gt 0) {
                            $nProjectsWithIncompleteFiles++
                            Write-LogWarning "  Skipping task custom field processing for $projectName — the following required file(s) were not found:"
                            foreach ($missingFile in $missingFiles) {
                                Write-LogWarning "    - $missingFile"
                            }
                        }
                        else {
                            Write-Host "Reading Reporting JSON file: $reportingTasksJsonFileName" -ForegroundColor Gray
                            try {
                                $reportingTasks = Get-Content -LiteralPath $reportingTasksJsonPath -Raw | ConvertFrom-Json
                                $publishedXml = [xml](Get-Content -LiteralPath $publishedXmlPath -Raw)
                                $publishedMppXml = [xml](Get-Content -LiteralPath $publishedMppXmlPath -Raw)

                                if ($reportingTasks -and $reportingTasks.PSObject.Properties.Name -contains 'ReportingProjectTasksData') {
                                    $rtd = $reportingTasks.ReportingProjectTasksData

                                    if ($rtd -and $rtd.PSObject.Properties.Name -contains 'Tasks') {
                                        $tasksCollectionNode = $rtd.Tasks

                                        foreach ($taskNode in $tasksCollectionNode) {
                                            $taskName = $taskNode.TaskName
                                            $taskGuid = $taskNode.TaskUId

                                            # Step 1: resolve the integer UID from the published XML by matching the task GUID.
                                            # PowerShell string -eq is case-insensitive, so no normalisation needed.
                                            $xmlTaskMatch = $publishedXml.Project.Tasks.Task |
                                            Where-Object { $_.GUID -eq $taskGuid } |
                                            Select-Object -First 1

                                            if ($null -eq $xmlTaskMatch) {
                                                Write-LogWarning "    Warning: GUID '$taskGuid' ($taskName) not found in published XML. Skipping."
                                                $nProjectTasksSkipped++
                                                continue
                                            }

                                            $taskIntUid = $xmlTaskMatch.UID

                                            # UID 0 is the project summary task — which is not published to Altus, skip silently.
                                            if ($taskIntUid -eq '0') { continue }

                                            # Step 2: resolve the Altus Task ID (Text30) from the mpp XML using the integer UID.
                                            $mppTaskMatch = $publishedMppXml.Project.Tasks.Task |
                                            Where-Object { $_.UID -eq $taskIntUid } |
                                            Select-Object -First 1

                                            if ($null -eq $mppTaskMatch -or [string]::IsNullOrEmpty($mppTaskMatch.Text30)) {
                                                Write-LogWarning "    Warning: No Altus Task ID (Text30) found for UID $taskIntUid ($taskName). Skipping."
                                                $nProjectTasksSkipped++
                                                continue
                                            }

                                            $altusTaskId = $mppTaskMatch.Text30
                                            Write-Log "  Task: $taskName | UID: $taskIntUid | Altus Task ID: $altusTaskId" Gray

                                            # then proceed with updating the altus Task record with the custom field mappings defined in Add-TaskDataverseFieldMappings
                                            try {
                                                $updateBody = Get-TaskDataverseBody -TaskGuid $altusTaskId -ReportingTask $taskNode
                                                Remove-BlocklistedTaskFields -TaskBody $updateBody

                                                # Extract many-to-many relationships before updating the record
                                                $manyToManyRelationships = Split-TaskManyToManyRelationships -Body $updateBody

                                                if ($updateBody.Count -gt 0) {
                                                    if ($ExecutionMode) {
                                                        Update-Record -setName 'sensei_tasks' -id $altusTaskId -body $updateBody
                                                        Write-Log "    Updated Task record in Dataverse for $taskName" Green
                                                    }
                                                    else {
                                                        Write-Log "    [What-If] Would update Task record in Dataverse for $taskName." Magenta
                                                    }
                                                    $nTasksUpdated++
                                                    $nProjectTasksUpdated++
                                                }
                                                else {
                                                    Write-Log "    No updates needed for Task $taskName" Gray
                                                    $nTasksNotRequiringUpdate++
                                                    $nProjectTasksSkipped++
                                                }

                                                Add-TaskManyToManyRelationships -TaskId $altusTaskId -ManyToManyRelationships $manyToManyRelationships
                                            }
                                            catch {
                                                $nErrored++
                                                $nProjectTasksErrored++
                                                Write-LogWarning "    Warning: Skipping task '$taskName' (Altus Task ID: $altusTaskId) — $_"
                                            }
                                        }
                                    }
                                }
                            }
                            catch {
                                $nErrored++
                                $nProjectTasksErrored++
                                Write-LogWarning "  Warning: Could not parse reporting JSON '$reportingTasksJsonFileName': $_"
                            }
                        }
                        $projectSummaries += [PSCustomObject]@{ Project = $projectName; Updated = $nProjectTasksUpdated; Skipped = $nProjectTasksSkipped; Errored = $nProjectTasksErrored }
                    } 
                }
                else {
                    Write-Host "No corresponding JSON file was found for: $jsonFileName. Skipping." -ForegroundColor Yellow
                    continue
                }
            }
            catch {
                $nErrored++
                Write-LogError "  ✗ Error processing $($mpp.Name): $_"
            }
        }
        
        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Per-Project Task Summary" -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        foreach ($ps in $projectSummaries) {
            Write-Host "  $($ps.Project): Updated: $($ps.Updated) | Skipped: $($ps.Skipped) | Errored: $($ps.Errored)" -ForegroundColor Cyan
        }

        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Import Task Fields Summary" -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Execution Mode:                 $executionModeText" -ForegroundColor Cyan
        Write-Host "Projects Processed:             $nProjectsIterated" -ForegroundColor Cyan
        Write-Host "Projects with Incomplete Files: $nProjectsWithIncompleteFiles" -ForegroundColor Cyan
        Write-Host "Tasks Updated:                  $nTasksUpdated" -ForegroundColor Cyan
        Write-Host "Tasks Not Requiring Update:     $nTasksNotRequiringUpdate" -ForegroundColor Cyan
        Write-Host "Errors Encountered:             $nErrored" -ForegroundColor Cyan
    }

    Write-Host "Import Task Fields completed" -ForegroundColor Green
}

<#
.SYNOPSIS
Retrieves all projects from Dataverse that have an external project ID.

.DESCRIPTION
Queries the Dataverse environment for all projects that have the sensei_externalprojectid 
field populated. This includes both projects with and without external project references. 
The function expands the external project relationship to retrieve external system information.

.OUTPUTS
Returns an array of project records containing:
- sensei_name: The project name
- sensei_externalprojectid: The external project identifier
- sensei_externalproject_project_sensei_pro: Expanded external project relationship data

.EXAMPLE
$projectsWithExternalId = Get-ProjectsWithExternalId
# Filter to projects from Project Desktop
$projectDesktopProjects = $projectsWithExternalId | Where-Object { 
    $_.sensei_externalproject_project_sensei_pro._sensei_externalsystem_value -eq $ProjectDesktopExternalSystemId 
}

.NOTES
This function is used to identify which projects have already been imported and may need 
updating rather than creating new records.
#>
function Get-ProjectsWithExternalId {
    Write-Host '--Retrieving Projects with External ID but no External Project Reference--'

    $projects = Get-AllRecords `
        -setName 'sensei_projects' `
        -query "?`$select=sensei_name,sensei_externalprojectid&`$filter=sensei_externalprojectid ne null&`$expand=sensei_externalproject_project_sensei_pro(`$select=sensei_name,_sensei_externalsystem_value)"

    return $projects
}

<#
.SYNOPSIS
Retrieves the version number of the Altus solution installed in Dataverse.

.DESCRIPTION
Queries the Dataverse solutions table to get the version of the SenseiProjectIndependent 
solution (Altus). This is used to verify the environment has the required Altus solution 
installed before proceeding with the import.

.OUTPUTS
Returns the version string of the Altus solution (e.g., "1.2.3.4").
Returns $null if the solution is not found.

.EXAMPLE
$version = Get-SolutionVersion
if ($version) {
    Write-Host "Altus Solution Version: $version"
}

.NOTES
The function looks for the solution with unique name 'SenseiProjectIndependent'.
If this solution is not found, the import process should not proceed.
#>
function Get-SolutionVersion {
    Write-Host '--Retrieving Altus Solution Version--'

    $solutionVersion = Get-AllRecords `
        -setName 'solutions' `
        -query "?`$select=version&`$filter=uniquename eq 'SenseiProjectIndependent'"

    return $solutionVersion.version
}

<#
.SYNOPSIS
Retrieves the organization name from the Dataverse environment.

.DESCRIPTION
Queries the Dataverse organizations table to get the name of the current organization.
This is primarily used for logging and display purposes to help identify which environment 
the import is running against.

.OUTPUTS
Returns the organization name as a string.

.EXAMPLE
$orgName = Get-OrgName
Write-Host "Organization: $orgName"

.NOTES
Every Dataverse environment has exactly one organization record.
#>
function Get-OrgName {
    Write-Host '--Retrieving Organization Info--'

    $orgInfo = Get-AllRecords `
        -setName 'organizations' `
        -query '?$select=organizationid,name'

    return $orgInfo.name
}

<#
.SYNOPSIS
Retrieves the Power Platform environment ID for the current Dataverse environment.

.DESCRIPTION
Gets the environment ID (GUID) for the connected Dataverse environment. The function first 
checks if an environment ID was manually configured in Defaults.ps1. If not configured, 
it queries the Dataverse.

The environment ID is required for creating external project references in Dataverse.

.OUTPUTS
Returns the environment ID as a string (GUID format).
Returns $null if the environment ID cannot be determined.

.EXAMPLE
$envId = Get-EnvironmentId
if ($envId) {
    Write-Host "Environment ID: $envId"
} else {
    Write-Host "Could not determine environment ID"
}

.NOTES
- If you encounter issues with automatic retrieval, you can manually configure the 
  environment ID in Defaults.ps1 by setting $global:EnvironmentId
- Automatic retrieval requires permissions to query the Dataverse
- The environment ID is typically a GUID format like "00000000-0000-0000-0000-000000000000"
#>
function Get-EnvironmentId {
    Write-Host '--Retrieving Environment ID--'

    # Check if manually configured in Defaults.ps1
    if ($global:EnvironmentId) {
        Write-Host "  Using configured Environment ID" -ForegroundColor Gray
        return $global:EnvironmentId
    }

    try {
        Write-Host "  Querying Dataverse for Environment ID" -ForegroundColor Gray

        # Get the environment ID
        $request = @{
            Uri     = $global:baseURI + "RetrieveCurrentOrganization(AccessType='Default')"
            Method  = 'Get'
            Headers = $global:baseHeaders
        }
        $orgResponse = Invoke-ResilientRestMethod -request $request
        $global:EnvironmentId = $orgResponse.Detail.EnvironmentId
        
        Write-Host "  Environment ID: $global:EnvironmentId" -ForegroundColor Gray
        return $global:EnvironmentId
    }
    catch {
        Write-Host "  Warning: Could not retrieve Environment ID: $_" -ForegroundColor Yellow
        return $null
    }
}



