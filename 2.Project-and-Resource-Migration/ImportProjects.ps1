# ============================================================================
# PROJECT IMPORT SCRIPT
# ============================================================================
# 
# USER CONFIGURATION: See "USER CONFIGURATION (EDIT HERE)" section below
#
# ============================================================================

. $PSScriptRoot\Defaults.ps1
. $PSScriptRoot\ProjectMappingHelpers.ps1

# Load Windows Forms for DoEvents support in COM automation
Add-Type -AssemblyName System.Windows.Forms

# ------------------------------
# USER CONFIGURATION (EDIT HERE)
# ------------------------------

<#
.SYNOPSIS
Maps Project Online project data to Dataverse project fields.

.DESCRIPTION
This function is the primary customisation point for mapping Project Online project fields 
to your Dataverse project table columns. Add your custom field mappings here to control 
how data flows from Project Online to Dataverse during the import process.

The function is called for each project being imported.

.PARAMETER Project
The hashtable representing the Dataverse project record being created/updated. 
Add field mappings to this hashtable using Dataverse logical column names.

.PARAMETER ProjectName
The name of the project.

.PARAMETER ProjectGuid
The unique GUID of the project from Project Online.

.PARAMETER ReportingProject
The Project Online reporting data for this project (if available). Contains OOTB fields 
and custom fields exported from Project Online. 

.EXAMPLE
# Map a text field from Project Online
$textValue = Get-ReportingCustomFieldTextValue -ReportingProject $ReportingProject -CustomFieldName 'Division'
if ($textValue) { $Project['cr62c_division'] = $textValue }

.EXAMPLE
# Map a lookup field
Set-ProjectLookupField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Department' -LookupName 'Department' -DataverseFieldName 'cr62c_Department'

.EXAMPLE
# Map a choice field
Set-ProjectChoiceField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Status' -ChoiceName 'ProjectStatus' -DataverseFieldName 'cr62c_status'

.NOTES
- Use Dataverse logical column names (lowercase) for field mappings
- All target columns must already exist in Dataverse
- Lookup and choice data must be pre-loaded via Defaults.ps1
- See the commented examples within the function for more mapping patterns
#>
function Add-ProjectDataverseFieldMappings {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Project,

        [Parameter(Mandatory = $true)]
        [string]$ProjectName,

        [Parameter(Mandatory = $true)]
        [string]$ProjectGuid,

        [Parameter(Mandatory = $false)]
        $ReportingProject
    )

    # Add Project -> Dataverse column mappings here.
    # Use Dataverse logical column names (lowercase). Columns must already exist.

    # --- OOTB fields (from $ReportingProject or $ProjectName) ---
    # $Project['cr_project_text_ootb'] = $ProjectName
    # $Project['cr_project_datetime_ootb'] = ($ReportingProject.ProjectStartDate -as [datetime])
    # $Project['cr_project_dateonly_ootb'] = ($ReportingProject.ProjectStartDate -as [datetime]).ToString("yyyy-MM-dd")
    # $Project['cr_project_whole_ootb'] = ($ReportingProject.ProjectIdentifier -as [int])
    # $Project['cr_project_decimal_ootb'] = ($ReportingProject.ProjectCalendarDuration -as [decimal])

    # --- Enterprise Custom Fields (from $ReportingProject.CustomFields[]) ---
    # $textValue = Get-ReportingCustomFieldTextValue -ReportingProject $ReportingProject -CustomFieldName 'Your Text Field'
    # if ($textValue) { $Project['cr_project_text_custom'] = [string]$textValue }
    #
    # $dateValue = Get-ReportingCustomFieldTextValue -ReportingProject $ReportingProject -CustomFieldName 'Your Date Field'
    # if ($dateValue) { $Project['cr_project_datetime_custom'] = ($dateValue -as [datetime]) }
    # if ($dateValue) { $Project['cr_project_dateonly_custom'] = ($dateValue -as [datetime]).ToString("yyyy-MM-dd") }

    # --- Lookup Field Example ---
    # Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file 
    #       so that the correct corresponding lookup value can be populated.
    # Also note that -DataverseFieldName is CASE SENSITIVE when identifying the Dataverse Schema Name for the lookup field
    # Set-ProjectLookupField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Organisational Change Impact' -LookupName 'ChangeImpact' -DataverseFieldName 'cr62c_OrganisationalChangeImpact'

    #--- Choice Field Example ---
    # Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding choice value can be populated.
    # If mapping to a multi-select choice field, set -IsMulti to $true
    # Set-ProjectChoiceField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Investment Category' -ChoiceName 'Investment Category' -DataverseFieldName 'sensei_investmentcategory'

    #--- Many-To-Many Relationship Example ---
    # Note: To map to a Dataverse many-to-many relationship, you must ensure the related table data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding related records can be populated.
    # This example demonstrates handling multi-select Project Online lookup fields that populate a Dataverse relationship table.
    # The relationship table is populated AFTER the project is created/updated using the Add-ToCollection function.
    # Note that the relationship table is populated additively only, so existing records in the relationship will be retained.
    # Set-ProjectManyToManyRelationship -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Multi Department' -LookupName 'Department' -RelationshipName 'cr62c_sensei_project_cr62c_Department_cr62c_Department' -RelatedTableCollectionName 'cr62c_departments' -RelatedTablePrimaryKey 'cr62c_departmentid'
}

<#
.SYNOPSIS
Creates the Dataverse project record body with default fields and custom mappings.

.DESCRIPTION
Helper function that constructs the complete hashtable for creating or updating a project 
in Dataverse. It initializes required OOTB fields (name, project type, external project ID) 
and then calls Add-ProjectDataverseFieldMappings to add any custom field mappings defined 
by the user.

This function is called internally by the import process and should not typically need 
to be modified.

.PARAMETER ProjectName
The name of the project.

.PARAMETER ProjectGuid
The unique GUID of the project from Project Online.

.PARAMETER DefaultProjectTypeId
The Dataverse ID of the default project type to assign to projects.

.PARAMETER ReportingProject
The Project Online reporting data for this project.

.OUTPUTS
Returns a hashtable containing all field mappings ready to be used in a Dataverse 
New-Record or Update-Record operation.

.EXAMPLE
$projectBody = Get-ProjectDataverseBody -ProjectName "Project Alpha" -ProjectGuid "12345678-1234-1234-1234-123456789012" -DefaultProjectTypeId "..." -ReportingProject $reportingData

.NOTES
- This function sets the required OOTB Altus fields
- Custom field mappings are added via Add-ProjectDataverseFieldMappings
#>
function Get-ProjectDataverseBody {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectName,

        [Parameter(Mandatory = $true)]
        [string]$ProjectGuid,

        [Parameter(Mandatory = $true)]
        [string]$DefaultProjectTypeId,

        [Parameter(Mandatory = $false)]
        $ReportingProject
    )

    $project = @{
        'sensei_name'                   = $ProjectName
        'sensei_projecttype@odata.bind' = "/sensei_enterpriseprojecttypes($DefaultProjectTypeId)"
        'sensei_externalprojectid'      = "ProjectDesktop_$ProjectGuid"
    }

    Add-ProjectDataverseFieldMappings -Project $project -ProjectName $ProjectName -ProjectGuid $ProjectGuid -ReportingProject $ReportingProject
    return $project
}

# ====== Don't edit below this line ======

<#
.SYNOPSIS
Imports projects into Dataverse.

.DESCRIPTION
The ImportProjects function handles the import of projects from JSON files into Dataverse.

.PARAMETER ExecutionMode
Controls whether to actually execute the import or run in test mode. When $true, records are created. When $false, runs in test mode. Default is $false.

.EXAMPLE
ImportProjects -ExecutionMode $true
This example imports projects into Dataverse, creating records.
#>

function ImportProjects {
    param(
        [Parameter(Mandatory = $false)]
        [bool]$ExecutionMode = $false
    )

    $executionModeText = if ($ExecutionMode) { "Execute" } else { "What-If" }
    Write-Host "Import Projects (ExecutionMode: $executionModeText)" -ForegroundColor Green    
    
    if ($EnableProjectNameFallback) {
        Write-Host "*** Enable Project Name Fallback mode is active ***" -ForegroundColor Green
    }

    Invoke-DataverseCommands {
        $nMPPsProcessed = 0
        $nProjectsCreated = 0
        $nProjectsUpdated = 0
        $nExternalProjectsCreated = 0
        $nMPPsUpdated = 0
        $nMPPsNotRequiringUpdate = 0
        $nErrored = 0

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
        } else {
            Write-Host "Altus Solution: Not found" -ForegroundColor Red
            return
        }

        #Read Organization info from environment
        $orgName = Get-OrgName
        if ($orgName) {
            Write-Host "Organization Name: $orgName"
        } else {
            Write-Host "Organization Name: Not found" -ForegroundColor Red
            return
        }

        #Get EnvironmentId for this environment
        $envId = Get-EnvironmentId
        if ($envId) {
            Write-Host "Environment ID: $envId"
        } else {
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
            $nMPPsProcessed++
            Write-Host "`n-- Processing MPP file: $($mpp.Name)..." -ForegroundColor Cyan
            
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

                    # --- Load Additional Metadata (Required for Custom Field Mapping) ---
                    # The _published.json has limited metadata. The _reporting.json contains project-level
                    # reporting fields (including custom fields) useful for mapping into Dataverse.
                    $reportingProject = $null
                    $reportingJsonFileName = $mpp.Name.Replace("_published.mpp", "_reporting.json")
                    $reportingJsonPath = Join-Path $filesPath $reportingJsonFileName
                    if (Test-Path -LiteralPath $reportingJsonPath) {
                        Write-Host "Reading Reporting JSON file: $reportingJsonFileName" -ForegroundColor Gray
                        try {
                            $reportingData = Get-Content -LiteralPath $reportingJsonPath -Raw | ConvertFrom-Json

                            if ($reportingData -and $reportingData.PSObject.Properties.Name -contains 'ReportingProjectData') {
                                $rpd = $reportingData.ReportingProjectData

                                if ($rpd -and $rpd.PSObject.Properties.Name -contains 'Project') {
                                    $projectNode = $rpd.Project
                                    if ($projectNode -is [System.Array]) {
                                        $reportingProject = $projectNode | Where-Object {
                                            ($_.ProjectUID -eq $projectGuid) -or ($_.ProjectUId -eq $projectGuid)
                                        } | Select-Object -First 1
                                    } else {
                                        if (($projectNode.ProjectUID -eq $projectGuid) -or ($projectNode.ProjectUId -eq $projectGuid)) {
                                            $reportingProject = $projectNode
                                        }
                                    }
                                } elseif ($rpd -is [System.Array]) {
                                    $reportingProject = $rpd | Where-Object {
                                        ($_.ProjectUID -eq $projectGuid) -or ($_.ProjectUId -eq $projectGuid)
                                    } | Select-Object -First 1
                                }
                            }

                            if (-not $reportingProject) {
                                Write-Host "  Warning: Reporting data loaded but ProjectUId '$projectGuid' not found in ReportingProjectData." -ForegroundColor Yellow
                            }
                        } catch {
                            Write-Host "  Warning: Could not parse reporting JSON '$reportingJsonFileName': $_" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Reporting JSON not found for custom fields: $reportingJsonFileName" -ForegroundColor DarkYellow
                    }
                    # --------------------------------------------------------

                    # Check if Project exists in Dataverse (based on sensei_externalprojectid match)
                    $existingProject = $projectDesktopProjects | Where-Object { $_.sensei_externalprojectid -eq "ProjectDesktop_$projectGuid" }
                    #check if there is a sensei_project record with the sensei_externalprojectid but no external project reference
                    $orphanedProject = $projectsWithExternalIdButNoExternalProjectReference | Where-Object { $_.sensei_externalprojectid -eq "ProjectDesktop_$projectGuid" }
                    if ($orphanedProject) {
                        Write-Host "Found orphaned Project record for $projectName ($projectGuid) in Dataverse." -ForegroundColor Yellow
                    }

                    $fallbackProjectNameMatch = $null
                    # Fallback to Project Name, (if enabled) to find existing project
                    if (-not $existingProject -and $EnableProjectNameFallback) {
                        Write-Host "No existing Project found by External ID. Attempting fallback by Project Name..." -ForegroundColor Yellow
                        $existingProjectByName = Get-ProjectsWithMatchingName -ProjectName $projectName
                        
                        if ($existingProjectByName.Count -eq 1) {
                            $fallbackProjectNameMatch = $existingProjectByName[0]
                            Write-Host "Found existing Project by exact Name match for $projectName ($projectGuid) in Dataverse." -ForegroundColor Yellow
                        } elseif ($existingProjectByName.Count -gt 1) {
                            Write-Host "ERROR: Multiple projects ($($existingProjectByName.Count)) found with name '$projectName' and no external project reference. Cannot proceed with this project." -ForegroundColor Red
                            $nErrored++
                            continue
                        } else {
                            Write-Host "No existing Project found by Name for $projectName ($projectGuid). Safe to proceed with creation of a new Project." -ForegroundColor Yellow
                        }
                    }

                    # If exists, update it
                    if ($existingProject) {
                        $nProjectsUpdated++
                        $newProjectId = $existingProject.sensei_projectid

                        if ($ExecutionMode) {
                            Write-Host "Updating existing Project record in Dataverse for $projectName..." -ForegroundColor Yellow
                            $updateBody = Get-ProjectDataverseBody -ProjectName $projectName -ProjectGuid $projectGuid -DefaultProjectTypeId $DefaultProjectTypeId -ReportingProject $reportingProject

                            # Extract many-to-many relationships before updating the record
                            $manyToManyRelationships = Split-ManyToManyRelationships -Body $updateBody

                            if ($updateBody.Count -gt 0) {
                                Update-Record -setName 'sensei_projects' -id $newProjectId -body $updateBody
                                Write-Host "Updated Project record in Dataverse for $projectName" -ForegroundColor Green
                            } else {
                                Write-Host "No updates needed for Project $projectName" -ForegroundColor Gray
                            }

                            # Process many-to-many relationships after update
                            Add-ProjectManyToManyRelationships -ProjectId $newProjectId -ManyToManyRelationships $manyToManyRelationships
                        } else {
                            Write-Host "[What-If] Would update existing Project record in Dataverse for $projectName." -ForegroundColor Magenta
                        }

                    } else {
                        # Create Project record in Dataverse
                        if ($ExecutionMode) {
                            if ($null -eq $orphanedProject -and $null -eq $fallbackProjectNameMatch) {
                                Write-Host "Creating Project record in Dataverse for $projectName..." -ForegroundColor Green

                                $project = Get-ProjectDataverseBody -ProjectName $projectName -ProjectGuid $projectGuid -DefaultProjectTypeId $DefaultProjectTypeId -ReportingProject $reportingProject

                                # Extract many-to-many relationships before creating the record
                                $manyToManyRelationships = Split-ManyToManyRelationships -Body $project

                                $newProjectId = New-Record -setName 'sensei_projects' -body $project
                                $nProjectsCreated++
                                Write-Host "Created Project record in Dataverse for $projectName with ID: $newProjectId" -ForegroundColor Green

                                # Process many-to-many relationships after project creation
                                Add-ProjectManyToManyRelationships -ProjectId $newProjectId -ManyToManyRelationships $manyToManyRelationships
                            } else {
                                if ($orphanedProject) {
                                    $newProjectId = $orphanedProject.sensei_projectid
                                    Write-Host "Using existing orphaned Project record in Dataverse for $projectName with ID: $newProjectId" -ForegroundColor Green
                                } elseif ($fallbackProjectNameMatch) {
                                    $newProjectId = $fallbackProjectNameMatch.sensei_projectid
                                    Write-Host "Using existing fallback Project record in Dataverse for $projectName with ID: $newProjectId" -ForegroundColor Green
                                }

                                $nProjectsUpdated++
                                $updateBody = Get-ProjectDataverseBody -ProjectName $projectName -ProjectGuid $projectGuid -DefaultProjectTypeId $DefaultProjectTypeId -ReportingProject $reportingProject
                                # Extract many-to-many relationships before updating the record
                                $manyToManyRelationships = Split-ManyToManyRelationships -Body $updateBody
                                if ($updateBody.Count -gt 0) {
                                    Write-Host "Updating Project fields..." -ForegroundColor Green
                                    Update-Record -setName 'sensei_projects' -id $newProjectId -body $updateBody
                                }

                                # Process many-to-many relationships after update
                                Add-ProjectManyToManyRelationships -ProjectId $newProjectId -ManyToManyRelationships $manyToManyRelationships
                            }

                            Write-Host "Creating External Project record in Dataverse for $projectName..." -ForegroundColor Green
                            $externalProject = @{
                                'sensei_name'                      = $projectName
                                'sensei_externalsystem@odata.bind' = "/sensei_externalsystems($ProjectDesktopExternalSystemId)"
                                'sensei_externalprojectidentifier' = "ProjectDesktop_$projectGuid"
                                'sensei_isprimary'                 = $true
                                'sensei_project@odata.bind'        = "/sensei_projects($newProjectId)"
                            }
                            $newExternalProjectId = New-Record -setName 'sensei_externalprojects' -body $externalProject
                            $nExternalProjectsCreated++
                            Write-Host "Created External Project record in Dataverse for $projectName with ID: $newExternalProjectId" -ForegroundColor Green
                        } else {
                            if ($null -eq $orphanedProject -and $null -eq $fallbackProjectNameMatch) {
                                Write-Host "[What-If] Would create Project record in Dataverse for $projectName." -ForegroundColor Magenta
                                $nProjectsCreated++
                                $newProjectId = [Guid]::NewGuid()
                            } else {
                                if ($orphanedProject) {
                                    $newProjectId = $orphanedProject.sensei_projectid
                                    Write-Host "[What-If] Would update existing orphaned Project record in Dataverse for $projectName with ID: $newProjectId" -ForegroundColor Magenta
                                } elseif ($fallbackProjectNameMatch) {
                                    $newProjectId = $fallbackProjectNameMatch.sensei_projectid
                                    Write-Host "[What-If] Would update existing Project with matching name in Dataverse for $projectName with ID: $newProjectId" -ForegroundColor Magenta
                                }
                                $nProjectsUpdated++
                            }
                            Write-Host "[What-If] Would create External Project record in Dataverse for $projectName." -ForegroundColor Magenta
                            #for the purposes of checking custom properties, we need a project id even in what-if mode
                            $nExternalProjectsCreated++
                        }
                    }

                    # Attempt to set custom properties using VBScript
                    Write-Host "Checking custom properties in $($mpp.Name) using VBScript..." -ForegroundColor Cyan
                    
                    # Kill any hung MS Project processes before starting (prevents issues from sleep/lock)
                    $msProjectProcesses = Get-Process -Name "WINPROJ" -ErrorAction SilentlyContinue
                    if ($msProjectProcesses) {
                        Write-Host "  Cleaning up existing MS Project processes..." -ForegroundColor Yellow
                        $msProjectProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                    }
                    
                    $vbsPath = Join-Path $PSScriptRoot "SetMppProperty.vbs"
                    $externalProjectId = "ProjectDesktop_$projectGuid"
                    
                    #only proceed if we have the full set of data to populate custom properties
                    if ($externalProjectId -and $newProjectId -and $projectName -and $solutionVersion -and $orgName -and $envId) {
                        if (Test-Path $vbsPath) {
                            try {
                                # Build arguments as individual strings - avoid Invoke-Expression
                                $arg1 = $mpp.FullName
                                $arg2 = "ProjectId"
                                $arg3 = $externalProjectId
                                $arg4 = "ConnectedProjectId"
                                $arg5 = if ($newProjectId) { $newProjectId.ToString() } else { " " }
                                $arg6 = "ConnectedProjectName"
                                $arg7 = if ($projectName) { $projectName.ToString() } else { " " }
                                $arg8 = "ConnectedProjectDescription"
                                $arg9 = " "  # Use space instead of empty string - PowerShell won't pass empty strings
                                $arg10 = "ConnectedEnvironmentSolutionVersion"
                                $arg11 = if ($solutionVersion) { $solutionVersion.ToString() } else { " " }
                                $arg12 = "ConnectedEnvironmentName"
                                $arg13 = if ($orgName) { $orgName.ToString() } else { " " }
                                $arg14 = "ConnectedConnectionId"
                                $arg15 = $envId.ToString()
                                $arg16 = if ($ExecutionMode) { "True" } else { "False" }
                                
                                # Call VBScript with timeout handling (5 minutes max)
                                $timeoutSeconds = 300
                                $tempOutputFile = [System.IO.Path]::GetTempFileName()
                                $tempErrorFile = [System.IO.Path]::GetTempFileName()
                                
                                try {
                                    # Build argument string with proper quoting for paths with spaces
                                    $vbsArgs = "//NoLogo `"$vbsPath`" `"$arg1`" `"$arg2`" `"$arg3`" `"$arg4`" `"$arg5`" `"$arg6`" `"$arg7`" `"$arg8`" `"$arg9`" `"$arg10`" `"$arg11`" `"$arg12`" `"$arg13`" `"$arg14`" `"$arg15`" `"$arg16`""
                                    
                                    $processInfo = Start-Process -FilePath "cscript.exe" `
                                        -ArgumentList $vbsArgs `
                                        -NoNewWindow `
                                        -PassThru `
                                        -RedirectStandardOutput $tempOutputFile `
                                        -RedirectStandardError $tempErrorFile
                                    
                                    # Wait with timeout
                                    $completed = $processInfo.WaitForExit($timeoutSeconds * 1000)
                                    
                                    if (-not $completed) {
                                        Write-Host "  ✗ VBScript timeout after $timeoutSeconds seconds - killing process" -ForegroundColor Red
                                        $processInfo.Kill()
                                        # Also kill any MS Project processes
                                        Get-Process -Name "WINPROJ" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                                        throw "VBScript execution timed out"
                                    }
                                    
                                    # Read output from both files
                                    $vbsOutput = @()
                                    if (Test-Path $tempOutputFile) {
                                        $vbsOutput += Get-Content $tempOutputFile -ErrorAction SilentlyContinue
                                    }
                                    if (Test-Path $tempErrorFile) {
                                        $vbsOutput += Get-Content $tempErrorFile -ErrorAction SilentlyContinue
                                    }
                                    
                                    # Display VBScript output
                                    if ($vbsOutput) {
                                        $vbsOutput | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
                                    }
                                    
                                    $exitCode = $processInfo.ExitCode
                                } catch {
                                    $nErrored++
                                } finally {
                                    # Cleanup temp files
                                    if (Test-Path $tempOutputFile) {
                                        Remove-Item $tempOutputFile -Force -ErrorAction SilentlyContinue
                                    }
                                    if (Test-Path $tempErrorFile) {
                                        Remove-Item $tempErrorFile -Force -ErrorAction SilentlyContinue
                                    }
                                }
                                
                                if ($exitCode -eq 0) {
                                    # Check if no changes were needed
                                    $noChangesMade = $vbsOutput -match 'No changes made, skipping save'
                                    
                                    if ($ExecutionMode) {
                                        if ($noChangesMade) {
                                            Write-Host "  No changes were needed in MPP file $($mpp.Name)." -ForegroundColor Green
                                            $nMPPsNotRequiringUpdate++
                                        } else {
                                            Write-Host "  ✓ Custom properties set successfully" -ForegroundColor Green
                                            $nMPPsUpdated++
                                        }
                                    } else {
                                        if ($noChangesMade) {
                                            Write-Host "[What-If] No changes required in MPP file $($mpp.Name)." -ForegroundColor Magenta
                                            $nMPPsNotRequiringUpdate++
                                        } else {
                                            Write-Host "[What-If] Would set custom properties in MPP file $($mpp.Name)." -ForegroundColor Magenta
                                            $nMPPsUpdated++
                                        }
                                    }
                                } else {
                                    $nErrored++
                                    Write-Host "  ✗ VBScript failed (exit code: $LASTEXITCODE)" -ForegroundColor Red
                                    Write-Host "  Manual step required: Set custom properties in MS Project" -ForegroundColor Yellow
                                }
                            } catch {
                                $nErrored++
                                Write-Host "  ✗ Error calling VBScript: $_" -ForegroundColor Red
                                Write-Host "  Manual step required: Set custom properties in MS Project" -ForegroundColor Yellow
                            }
                        } else {
                            $nErrored++
                            Write-Host "  ✗ VBScript file not found: $vbsPath" -ForegroundColor Red
                            Write-Host "  Manual step required: Set custom properties in MS Project" -ForegroundColor Yellow
                        }
                    } else {
                        $nErrored++
                        Write-Host "  ✗ Insufficient data to set custom properties in MPP file." -ForegroundColor Red
                        Write-Host "    Ensure External Project ID, Project ID, Project Name, Solution Version, Organization Name, and Environment ID are all available." -ForegroundColor Yellow
                        Write-Host "  Manual step required: Set custom properties in MS Project" -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No corresponding JSON file was found for: $jsonFileName. Skipping." -ForegroundColor Yellow
                    continue
                }
            } catch {
                $nErrored++
                Write-Host "  ✗ Error processing $($mpp.Name): $_" -ForegroundColor Red
            }
        }
        
        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Import Projects Summary" -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Execution Mode:                $executionModeText" -ForegroundColor Cyan
        Write-Host "MPPs Processed:                $nMPPsProcessed" -ForegroundColor Cyan
        Write-Host "Projects Created:              $nProjectsCreated" -ForegroundColor Green
        Write-Host "Projects Updated:              $nProjectsUpdated" -ForegroundColor Yellow
        Write-Host "External Projects Created:     $nExternalProjectsCreated" -ForegroundColor Cyan
        Write-Host "MPPs Updated:                  $nMPPsUpdated" -ForegroundColor Cyan
        Write-Host "MPPs Not Requiring Update:     $nMPPsNotRequiringUpdate" -ForegroundColor Cyan
        Write-Host "Errors Encountered:            $nErrored" -ForegroundColor Cyan
    }

    Write-Host "Import Projects completed" -ForegroundColor Green
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
Retrieves a project from Dataverse that has an exact match on the name parameter.

.DESCRIPTION
Queries the Dataverse environment for projects that have an exact match on the Name.

.OUTPUTS
Returns an array of project records containing:
- sensei_name: The project name
- sensei_externalprojectid: The external project identifier
- sensei_externalproject_project_sensei_pro: Expanded external project relationship data

.EXAMPLE
$projectsWithMatchingName = Get-ProjectsWithMatchingName -ProjectName "Project Alpha"

.NOTES
This function is used to identify if there are existing projects in the environment
with an exact match on the Name. Will only return a project if there is a single match
with no external project ID or external project reference.
#>
function Get-ProjectsWithMatchingName {
    param (
        [string]$ProjectName
    )
    Write-Host '--Attempting to find Projects with Matching Name--'

    # Escape single quotes for OData literal, then URL-encode for query string safety (e.g. '&')
    $escapedProjectName = $ProjectName.Replace("'", "''")
    $urlSafeProjectName = [System.Uri]::EscapeDataString($escapedProjectName)

    $projects = Get-AllRecords `
        -setName 'sensei_projects' `
        -query "?`$select=sensei_name,sensei_externalprojectid&`$filter=sensei_name eq '$urlSafeProjectName'&`$expand=sensei_externalproject_project_sensei_pro(`$select=sensei_name,_sensei_externalsystem_value)"

    # Filter client-side to ensure exact match (OData may not handle trailing whitespace correctly)
    $projects = $projects | Where-Object { $_.sensei_name -eq $ProjectName }

    Write-Host "Found $($projects.Count) projects with direct match on name '$ProjectName'."

    if ($projects.Count -gt 0) {
        # Filter to only projects where the external project ID is not populated [we don't want to hijack existing external project]
        $projects = $projects | Where-Object {
            $null -eq $_.sensei_externalprojectid
        }
        Write-Host "$($projects.Count) projects remain after filtering out those with existing external project IDs."

        if ($projects.Count -gt 0) {
            $projects = $projects | Where-Object {
                $null -eq $_.sensei_externalproject_project_sensei_pro -or
                $_.sensei_externalproject_project_sensei_pro.Count -eq 0
            }

            if ($projects.Count -gt 0) {
                Write-Host "$($projects.Count) projects remain after filtering out those with existing external project references."
            } else {
                Write-Host "No projects with name '$ProjectName' and no external project reference found."
            }
        } else {
            Write-Host "No projects with name '$ProjectName' and no external project ID found."
        }
    } else {
        Write-Host "No projects found with name '$ProjectName'."
    }

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
    } catch {
        Write-Host "  Warning: Could not retrieve Environment ID: $_" -ForegroundColor Yellow
        return $null
    }
}
