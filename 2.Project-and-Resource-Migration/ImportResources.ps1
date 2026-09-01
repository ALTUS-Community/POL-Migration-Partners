# ============================================================================
# RESOURCE IMPORT SCRIPT
# ============================================================================
# 
# USER CONFIGURATION: See "USER CONFIGURATION (EDIT HERE)" section below
#
# ============================================================================

. $PSScriptRoot\ExtendedTableOperations.ps1
. $PSScriptRoot\Defaults.ps1
. $PSScriptRoot\ResourceMappingHelpers.ps1

# ------------------------------
# USER CONFIGURATION (EDIT HERE)
# ------------------------------

<#
.SYNOPSIS
Maps Project Online named resource data to Altus bookable resource fields.

.DESCRIPTION
This function is the primary customization point for mapping Project Online named resource 
fields to your Dataverse bookable resource table columns. Add your custom field mappings 
here to control how data flows from Project Online to Dataverse during the resource import process.

Named resources are those with a valid system user account.

The function is called for each named resource being imported and receives the Project Online 
resource data including OOTB fields and custom fields.

.PARAMETER ResourceBody
The hashtable representing the Dataverse bookable resource record being created/updated. 
Add field mappings to this hashtable using Dataverse logical column names.

.PARAMETER ProjectResource
The Project Online resource data for this resource. Contains OOTB fields (ResourceName, 
ResourceType, ResourceStandardRate, etc.) and custom fields exported from Project Online.

.EXAMPLE
# Map a text field from Project Online
$rbsValue = Get-ResourceCustomFieldTextValue -ProjectResource $ProjectResource -CustomFieldName 'RBS'
if ($rbsValue) { $ResourceBody['cr62c_rbs'] = $rbsValue }

.EXAMPLE
# Map a lookup field
Set-ResourceLookupField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'Department' -LookupName 'Department' -DataverseFieldName 'cr62c_Department'

.EXAMPLE
# Map a choice field
Set-ResourceChoiceField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'ResourceType' -ChoiceName 'ResourceType' -DataverseFieldName 'cr62c_resourcetype'

.NOTES
- Use Dataverse logical column names (lowercase) for field mappings
- All target columns must already exist in Dataverse
- Lookup and choice data must be pre-loaded via Defaults.ps1
- See the commented examples within the function for more mapping patterns
#>
function Add-NamedResourceDataverseFieldMappings {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ResourceBody,

        [Parameter(Mandatory = $true)]
        $ProjectResource
    )

    # Add Named Resource -> Dataverse column mappings here.
    # Use Dataverse logical column names (lowercase). Columns must already exist.

    # --- OOTB fields (from $ProjectResource) ---
    # $ResourceBody['cr_resource_text_ootb'] = $ProjectResource.ResourceName
    # $ResourceBody['cr_resource_datetime_ootb'] = ($ProjectResource.ResourceCreatedDate -as [datetime])
    # $ResourceBody['cr_resource_dateonly_ootb'] = ($ProjectResource.ResourceCreatedDate -as [datetime]).ToString("yyyy-MM-dd")
    # $ResourceBody['cr_resource_whole_ootb'] = ($ProjectResource.ResourceType -as [int])
    # $ResourceBody['cr_resource_decimal_ootb'] = ($ProjectResource.ResourceStandardRate -as [decimal])

    # --- Enterprise Custom Fields (from $ProjectResource.CustomFields[]) ---
    # $rbsValue = Get-ResourceCustomFieldTextValue -ProjectResource $ProjectResource -CustomFieldName 'RBS'
    # if ($rbsValue) { $ResourceBody['cr_resource_rbs'] = [string]$rbsValue }
    #
    # $dateValue = Get-ResourceCustomFieldTextValue -ProjectResource $ProjectResource -CustomFieldName 'Your Date Field'
    # if ($dateValue) { $ResourceBody['cr_resource_datetime_custom'] = ($dateValue -as [datetime]) }
    # if ($dateValue) { $ResourceBody['cr_resource_dateonly_custom'] = ($dateValue -as [datetime]).ToString("yyyy-MM-dd") }
    
    # --- Lookup Field Example ---
    # Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file 
    #       so that the correct corresponding lookup value can be populated.
    # Also note that -DataverseFieldName is CASE SENSITIVE when identifying the Dataverse Schema Name for the lookup field
    # Set-ResourceLookupField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'RBS' -LookupName 'RBS' -DataverseFieldName 'cr62c_RBSHierarchy'

    # --- Choice Field Example ---
    # Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding choice value can be populated.
    # Set-ResourceChoiceField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'Cost Type' -ChoiceName 'Cost Type' -DataverseFieldName 'cr62c_costtype'

    # ---Many-To-Many Relationship Lookup Example ---
    # Note: To map to a Dataverse many-to-many relationship table, you must ensure the related table's data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding values can be populated.
    # In order to populate the many-to-many relationship table, you must identify the some key aspects of the relationship table including;
    # -RelationshipName - the name of the relationship as defined in Dataverse
    # -RelatedTableCollectionName - the table Collection Name (plural) of the related table
    # -RelatedTablePrimaryKey - the primary key (id) field of the related table
    # Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
    # Set-ResourceManyToManyRelationship -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'Multi Skills' -LookupName 'Skills' -RelationshipName 'cr62c_sensei_bookableresource_cr62c_Skills_cr62c_Skills' -RelatedTableCollectionName 'cr62c_skillses' -RelatedTablePrimaryKey 'cr62c_skillsid'    
}

<#
.SYNOPSIS
Maps Project Online generic resource data to Altus bookable resource fields.

.DESCRIPTION
This function is the primary customization point for mapping Project Online generic resource 
fields to your Dataverse bookable resource table columns. Add your custom field mappings 
here to control how data flows from Project Online to Dataverse during the generic resource 
import process.

The function is called for each generic resource being imported and receives the Project Online 
resource data including OOTB fields and custom fields.

.PARAMETER ResourceBody
The hashtable representing the Dataverse bookable resource record being created/updated. 
Add field mappings to this hashtable using Dataverse logical column names.

.PARAMETER GenericResource
The Project Online generic resource data. Contains OOTB fields (e.g. ResourceName) and custom fields exported from Project Online.

.EXAMPLE
# Map a text field from Project Online
$skillValue = Get-ResourceCustomFieldTextValue -ProjectResource $GenericResource -CustomFieldName 'Primary Skill'
if ($skillValue) { $ResourceBody['cr62c_primaryskill'] = $skillValue }

.EXAMPLE
# Map a lookup field
Set-ResourceLookupField -ResourceBody $ResourceBody -ProjectResource $GenericResource -POLCustomFieldName 'Resource Category' -LookupName 'ResourceCategory' -DataverseFieldName 'cr62c_ResourceCategory'

.EXAMPLE
# Map a choice field
Set-ResourceChoiceField -ResourceBody $ResourceBody -ProjectResource $GenericResource -POLCustomFieldName 'Availability' -ChoiceName 'Availability' -DataverseFieldName 'cr62c_availability'

.NOTES
- Use Dataverse logical column names (lowercase) for field mappings
- All target columns must already exist in Dataverse
- Lookup and choice data must be pre-loaded via Defaults.ps1
- Generic resources are only imported when Mode is set to 'NamedAndGeneric'
- See the commented examples within the function for more mapping patterns
#>
function Add-GenericResourceDataverseFieldMappings {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ResourceBody,

        [Parameter(Mandatory = $true)]
        $GenericResource
    )

    # Add Generic Resource -> Dataverse column mappings here.
    # Use Dataverse logical column names (lowercase). Columns must already exist.

    # --- OOTB fields (from $GenericResource) ---
    # $ResourceBody['cr_generic_resource_text_ootb'] = $GenericResource.ResourceName
    # $ResourceBody['cr_generic_resource_datetime_ootb'] = ($GenericResource.ResourceCreatedDate -as [datetime])
    # $ResourceBody['cr_generic_resource_dateonly_ootb'] = ($GenericResource.ResourceCreatedDate -as [datetime]).ToString("yyyy-MM-dd")
    # $ResourceBody['cr_generic_resource_whole_ootb'] = ($GenericResource.ResourceType -as [int])
    # $ResourceBody['cr_generic_resource_decimal_ootb'] = ($GenericResource.ResourceStandardRate -as [decimal])

    # --- Enterprise Custom Fields (from $GenericResource.CustomFields[]) ---
    # $rbsValue = Get-ResourceCustomFieldTextValue -ProjectResource $GenericResource -CustomFieldName 'RBS'
    # if ($rbsValue) { $ResourceBody['cr_generic_resource_rbs'] = [string]$rbsValue }
    #
    # $dateValue = Get-ResourceCustomFieldTextValue -ProjectResource $GenericResource -CustomFieldName 'Your Date Field'
    # if ($dateValue) { $ResourceBody['cr_generic_resource_datetime_custom'] = ($dateValue -as [datetime]) }
    # if ($dateValue) { $ResourceBody['cr_generic_resource_dateonly_custom'] = ($dateValue -as [datetime]).ToString("yyyy-MM-dd") }

    # --- Lookup Field Example ---
    # Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file 
    #       so that the correct corresponding lookup value can be populated.
    # Also note that -DataverseFieldName is CASE SENSITIVE when identifying the Dataverse Schema Name for the lookup field
    # Set-ResourceLookupField -ResourceBody $ResourceBody -ProjectResource $GenericResource -POLCustomFieldName 'RBS' -LookupName 'RBS' -DataverseFieldName 'cr62c_RBSHierarchy'

    # --- Choice Field Example ---
    # Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file
    #       so that the correct corresponding choice value can be populated.
    # Set-ResourceChoiceField -ResourceBody $ResourceBody -ProjectResource $GenericResource -POLCustomFieldName 'Cost Type' -ChoiceName 'Cost Type' -DataverseFieldName 'cr62c_costtype'

    # ---Many-To-Many Relationship Lookup Example ---
    # Note: To map to a Dataverse many-to-many relationship table, you must ensure the related table's data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding values can be populated.
    # In order to populate the many-to-many relationship table, you must identify the some key aspects of the relationship table including;
    # -RelationshipName - the name of the relationship as defined in Dataverse
    # -RelatedTableCollectionName - the table Collection Name (plural) of the related table
    # -RelatedTablePrimaryKey - the primary key (id) field of the related table
    # Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
    # Set-ResourceManyToManyRelationship -ResourceBody $ResourceBody -ProjectResource $GenericResource -POLCustomFieldName 'Multi Skills' -LookupName 'Skills' -RelationshipName 'cr62c_sensei_bookableresource_cr62c_Skills_cr62c_Skills' -RelatedTableCollectionName 'cr62c_skillses' -RelatedTablePrimaryKey 'cr62c_skillsid'    
}

<#
.SYNOPSIS
Creates the Dataverse named bookable resource record body with default fields and custom mappings.

.DESCRIPTION
Helper function that constructs the complete hashtable for creating or updating a named 
bookable resource in Dataverse. It initializes required OOTB fields (name, resource type, 
system user reference, target utilization, primary role, enterprise calendar) and then calls 
Add-NamedResourceDataverseFieldMappings to add any custom field mappings defined by the user.

Named resources are linked to system users and have ResourceType = 955000001.

This function is called internally by the import process and should not typically need 
to be modified.

.PARAMETER ProjectResource
The Project Online resource data including OOTB fields and custom fields.

.PARAMETER MatchingSystemUser
The Dataverse system user record that matches this resource's ResourceNTAccount.

.PARAMETER DefaultTargetUtilisation
The default target utilization percentage to assign to the resource.

.PARAMETER DefaultPrimaryRoleId
The Dataverse ID of the default primary role (bookable resource) to assign.

.PARAMETER DefaultEnterpriseCalendarId
The Dataverse ID of the default enterprise calendar to assign.

.OUTPUTS
Returns a hashtable containing all field mappings ready to be used in a Dataverse 
New-Record or Update-Record operation.

.EXAMPLE
$resourceBody = New-NamedBookableResourceDataverseBody -ProjectResource $resource -MatchingSystemUser $user -DefaultTargetUtilisation 100 -DefaultPrimaryRoleId "..." -DefaultEnterpriseCalendarId "..."

.NOTES
- This function sets the required OOTB Altus bookable resource fields
- Custom field mappings are added via Add-NamedResourceDataverseFieldMappings
- Named resources must have a matching system user to be imported
#>
function New-NamedBookableResourceDataverseBody {
    param(
        [Parameter(Mandatory = $true)]
        $ProjectResource,

        [Parameter(Mandatory = $true)]
        $MatchingSystemUser,

        [Parameter(Mandatory = $true)]
        $DefaultTargetUtilisation,

        [Parameter(Mandatory = $true)]
        [string]$DefaultPrimaryRoleId,

        [Parameter(Mandatory = $true)]
        [string]$DefaultEnterpriseCalendarId
    )

    $resourceBody = @{
        'sensei_name'                          = $ProjectResource.ResourceName
        'sensei_resourcetype'                  = 955000001  # Named Resource
        'sensei_user@odata.bind'               = "/systemusers($($MatchingSystemUser.systemuserid))"
        'sensei_targetutilization'             = $DefaultTargetUtilisation
        'sensei_primaryrole@odata.bind'        = "/sensei_bookableresources($DefaultPrimaryRoleId)"
        'sensei_enterprisecalendar@odata.bind' = "/sensei_enterprisecalendars($DefaultEnterpriseCalendarId)"
    }

    Add-NamedResourceDataverseFieldMappings -ResourceBody $resourceBody -ProjectResource $ProjectResource
    return $resourceBody
}

<#
.SYNOPSIS
Creates the Dataverse generic bookable resource record body with default fields and custom mappings.

.DESCRIPTION
Helper function that constructs the complete hashtable for creating or updating a generic 
bookable resource in Dataverse. It initializes required OOTB fields (name and resource type) 
and then calls Add-GenericResourceDataverseFieldMappings to add any custom field mappings 
defined by the user.

Generic resources are not linked to system users and have ResourceType = 955000000.

This function is called internally by the import process and should not typically need 
to be modified.

.PARAMETER GenericResource
The Project Online generic resource data including OOTB fields and custom fields.

.OUTPUTS
Returns a hashtable containing all field mappings ready to be used in a Dataverse 
New-Record or Update-Record operation.

.EXAMPLE
$resourceBody = New-GenericBookableResourceDataverseBody -GenericResource $genericResource

.NOTES
- This function sets the required OOTB Altus bookable resource fields for generic resources
- Custom field mappings are added via Add-GenericResourceDataverseFieldMappings
- Generic resources are only imported when Mode is set to 'NamedAndGeneric'
- Unlike named resources, generic resources do not require system user references
#>
function New-GenericBookableResourceDataverseBody {
    param(
        [Parameter(Mandatory = $true)]
        $GenericResource
    )

    $resourceBody = @{
        'sensei_name'         = $GenericResource.ResourceName
        'sensei_resourcetype' = 955000000  # Generic Resource
    }

    Add-GenericResourceDataverseFieldMappings -ResourceBody $resourceBody -GenericResource $GenericResource
    return $resourceBody
}

# ====== Don't edit below this line ======

<#
.SYNOPSIS
Imports resources into Dataverse.

.DESCRIPTION
The ImportResources function handles the import of resources from JSON files into Dataverse.

.PARAMETER Mode
The type of resources to import. Valid values are 'NamedOnly' or 'NamedAndGeneric'. Default is 'NamedOnly'.

.PARAMETER ExecutionMode
Controls whether to actually execute the import or run in test mode. When $true, records are created. When $false, runs in test mode. Default is $false.

.EXAMPLE
ImportResources -Mode 'NamedOnly' -ExecutionMode $false
This example imports only named resources into Altus in test mode (no actual writes).

.EXAMPLE
ImportResources -Mode 'NamedAndGeneric' -ExecutionMode $true
This example imports named and generic resources into Altus and actually creates the records.
#>

function ImportResources {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('NamedOnly', 'NamedAndGeneric')]
        [string]$Mode = 'NamedOnly',
        
        [Parameter(Mandatory = $false)]
        [bool]$ExecutionMode = $false
    )
    
    $executionModeText = if ($ExecutionMode) { "Execute" } else { "What-If" }
    Write-Host "Import Resources (Mode: $Mode, ExecutionMode: $executionModeText)" -ForegroundColor Green
    
    Invoke-DataverseCommands {
        $nProjectOnlineFiles = 0
        $nProjectOnlineResources = 0
        $nProjectOnlineUniqueNamedResources = 0
        $nProjectOnlineUniqueGenericResources = 0
        $nNoMatchingSystemUsers = 0
        $nNamedBookableResourcesCreated = 0
        $nNamedBookableResourcesUpdated = 0
        $nGenericBookableResourcesCreated = 0
        $nGenericBookableResourcesUpdated = 0

        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Reading Bookable Resources..." -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan

        # Read all Bookable Resources from the Dataverse environment
        $bookableResources = Get-BookableResources
        Write-Host "Retrieved $($bookableResources.Count) Named Bookable Resources from Dataverse." 
        
        # Flatten the structure - copy domainname from nested sensei_user to top level
        foreach ($resource in $bookableResources) {
            if ($resource.sensei_user) {
                $resource | Add-Member -NotePropertyName 'domainname' -NotePropertyValue $resource.sensei_user.domainname -Force
            }
        }
        #        $bookableResources | Format-Table -AutoSize | Out-String | Write-Host
        
        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Reading System Users..." -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        # Read all System Users from the Dataverse environment
        $systemUsers = Get-SystemUsers
        Write-Host "Retrieved $($systemUsers.Count) System Users from Dataverse."
        #        $systemUsers | Format-Table -AutoSize | Out-String | Write-Host
        
        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Reading Project Online Resource Data..." -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        # Read exported Project Online resource files from the Files directory
        $filesPath = Join-Path $PSScriptRoot "Files"
        if (-not (Test-Path $filesPath)) {
            New-Item -Path $filesPath -ItemType Directory -Force | Out-Null
        }
        $resourceFiles = Get-ChildItem -Path $filesPath -Filter "Project_*_reporting_Resources.json"
        $nProjectOnlineFiles = $resourceFiles.Count
        Write-Host "Found $($resourceFiles.Count) Project Online resource files."
        
        $allProjectResources = @()
        foreach ($file in $resourceFiles) {
            Write-Host "Reading file: $($file.Name)" -ForegroundColor Gray
            $fileContent = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            $allProjectResources += $fileContent.ReportingProjectResourcesData.Resources
        }
        Write-Host "Loaded $($allProjectResources.Count) resources from Project Online files."
        $nProjectOnlineResources = $allProjectResources.Count

        # Deduplicate Project Resources based on ResourceUID
        $allProjectResources = $allProjectResources | Sort-Object -Property ResourceUID -Unique
        #Filter to only include ResourceType = 2 (Named Resources) and 20 (Generic Resources) [will only use Generic if that mode is selected]
        $allProjectResources = $allProjectResources | Where-Object { $_.ResourceType -eq 2 -or $_.ResourceType -eq 20 }
        # Sort by ResourceName
        $allProjectResources = $allProjectResources | Sort-Object -Property ResourceName
        
        $allNamedResources = $allProjectResources | Where-Object { $null -ne $_.ResourceNTAccount -and $_.ResourceNTAccount.Trim() -ne '' -and $_.ResourceType -eq 2 }
        $allGenericResources = $allProjectResources | Where-Object { $null -eq $_.ResourceNTAccount -or $_.ResourceNTAccount.Trim() -eq '' -and $_.ResourceType -eq 20 }        

        # Extract email from ResourceNTAccount (format: 'i:0#.f|membership|email@domain.com')
        foreach ($resource in $allNamedResources) {
            $ntAccountParts = $resource.ResourceNTAccount.Split('|')
            if ($ntAccountParts.Count -ge 3) {
                $resource | Add-Member -NotePropertyName 'ExtractedEmail' -NotePropertyValue $ntAccountParts[2] -Force
            }
        }
        
        # Remove any where we couldn't extract a login
        $allNamedResources = $allNamedResources | Where-Object { $null -ne $_.ExtractedEmail -and $_.ExtractedEmail.Trim() -ne '' }
        Write-Host "After deduplication, $($allNamedResources.Count) unique Named resources remain." 
        $nProjectOnlineUniqueNamedResources = $allNamedResources.Count

        if ($Mode -eq 'NamedAndGeneric') {
            Write-Host "After deduplication, $($allGenericResources.Count) unique Generic resources remain."
            $nProjectOnlineUniqueGenericResources = $allGenericResources.Count
        }

        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Processing Named Resources..." -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan

        foreach ($projectResource in $allNamedResources) {
            Write-Host "`n-- Processing Resource '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'..." -ForegroundColor Cyan
            
            # Check if there is an existing matching Bookable Resource
            $matchingBookableResource = $bookableResources | Where-Object { 
                $_.domainname -and $projectResource.ExtractedEmail -and 
                $_.domainname.ToLower() -eq $projectResource.ExtractedEmail.ToLower()
            }
            
            # If there is no existing matching System User, log this and skip to the next Resource
            $matchingSystemUser = $systemUsers | Where-Object { 
                $_.domainname -and $projectResource.ExtractedEmail -and 
                $_.domainname.ToLower() -eq $projectResource.ExtractedEmail.ToLower()
            }
            if (-not $matchingSystemUser) {
                $nNoMatchingSystemUsers++
                Write-Host "No matching System User found for Resource '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'. Skipping." -ForegroundColor DarkYellow
                continue
            }
            
            # Prepare the resource body
            $resourceBody = New-NamedBookableResourceDataverseBody `
                -ProjectResource $projectResource `
                -MatchingSystemUser $matchingSystemUser `
                -DefaultTargetUtilisation $DefaultTargetUtilisation `
                -DefaultPrimaryRoleId $DefaultPrimaryRoleId `
                -DefaultEnterpriseCalendarId $DefaultEnterpriseCalendarId
            
            if ($matchingBookableResource) {
                # Update existing resource
                $nNamedBookableResourcesUpdated++
                if (-not $ExecutionMode) {
                    Write-Host "[What-If] Would update existing Bookable Resource for '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'" -ForegroundColor Magenta
                } else {
                    Write-Host "Updating existing Bookable Resource for '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'." -ForegroundColor Yellow
                    
                    # Extract many-to-many relationships before updating the record
                    $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
                    
                    Update-Record -setName 'sensei_bookableresources' -id $matchingBookableResource.sensei_bookableresourceid -body $resourceBody
                    Write-Host "Updated Named Bookable Resource '$($projectResource.ResourceName)'." -ForegroundColor Green
                    
                    # Process many-to-many relationships after update
                    Add-ResourceManyToManyRelationships -ResourceId $matchingBookableResource.sensei_bookableresourceid -ManyToManyRelationships $manyToManyRelationships
                }
            } else {
                # Create new resource
                $nNamedBookableResourcesCreated++
                if (-not $ExecutionMode) {
                    Write-Host "[What-If] Would create Bookable Resource for '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'" -ForegroundColor Magenta
                } else {
                    Write-Host "Creating Bookable Resource for '$($projectResource.ResourceName)' with login '$($projectResource.ExtractedEmail)'." -ForegroundColor Green
                    
                    # Extract many-to-many relationships before creating the record
                    $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
                    
                    $newResource = New-Record -setName 'sensei_bookableresources' -body $resourceBody
                    Write-Host "Created Named Bookable Resource '$($projectResource.ResourceName)'." -ForegroundColor Green
                    
                    # Process many-to-many relationships after creation
                    Add-ResourceManyToManyRelationships -ResourceId $newResource -ManyToManyRelationships $manyToManyRelationships
                }
            }
        }

        if ($Mode -eq 'NamedAndGeneric') {
            
            Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
            Write-Host "Processing Generic Resources..." -ForegroundColor Cyan
            Write-Host "--------------------------------------------" -ForegroundColor Cyan

            foreach ($genericResource in $allGenericResources) {
                Write-Host "`n-- Processing Generic Resource '$($genericResource.ResourceName)'..." -ForegroundColor Cyan
                
                # Check if there is an existing matching Bookable Resource
                $matchingBookableResource = $bookableResources | Where-Object { 
                    $_.sensei_name -and $genericResource.ResourceName -and 
                    $_.sensei_name.ToLower() -eq $genericResource.ResourceName.ToLower()
                }
                
                # Prepare the resource body
                $resourceBody = New-GenericBookableResourceDataverseBody -GenericResource $genericResource
                
                if ($matchingBookableResource) {
                    # Update existing resource
                    $nGenericBookableResourcesUpdated++
                    if (-not $ExecutionMode) {
                        Write-Host "[What-If] Would update existing Generic Bookable Resource for '$($genericResource.ResourceName)'" -ForegroundColor Magenta
                    } else {
                        Write-Host "Updating existing Generic Bookable Resource for '$($genericResource.ResourceName)'." -ForegroundColor Yellow
                        
                        # Extract many-to-many relationships before updating the record
                        $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
                        
                        Update-Record -setName 'sensei_bookableresources' -id $matchingBookableResource.sensei_bookableresourceid -body $resourceBody
                        Write-Host "Updated Generic Bookable Resource '$($genericResource.ResourceName)'." -ForegroundColor Green
                        
                        # Process many-to-many relationships after update
                        Add-ResourceManyToManyRelationships -ResourceId $matchingBookableResource.sensei_bookableresourceid -ManyToManyRelationships $manyToManyRelationships
                    }
                } else {
                    # Create new resource
                    $nGenericBookableResourcesCreated++
                    if (-not $ExecutionMode) {
                        Write-Host "[What-If] Would create Generic Bookable Resource for '$($genericResource.ResourceName)'" -ForegroundColor Magenta
                    } else {
                        Write-Host "Creating Generic Bookable Resource for '$($genericResource.ResourceName)'." -ForegroundColor Green
                        
                        # Extract many-to-many relationships before creating the record
                        $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
                        
                        $newGenericResource = New-Record -setName 'sensei_bookableresources' -body $resourceBody
                        Write-Host "Created Generic Bookable Resource '$($genericResource.ResourceName)'." -ForegroundColor Green
                        
                        # Process many-to-many relationships after creation
                        Add-ResourceManyToManyRelationships -ResourceId $newGenericResource -ManyToManyRelationships $manyToManyRelationships
                    }
                }
            }
        }

        Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Import Resources Summary" -ForegroundColor Cyan
        Write-Host "--------------------------------------------" -ForegroundColor Cyan
        Write-Host "Execution Mode:                                        $executionModeText" -ForegroundColor Cyan
        Write-Host "Project Online Resource Files Processed:               $nProjectOnlineFiles" -ForegroundColor Cyan
        Write-Host "Project Online Resources Processed:                    $nProjectOnlineResources" -ForegroundColor Cyan
        Write-Host "Unique Named Resources Identified:                     $nProjectOnlineUniqueNamedResources" -ForegroundColor Cyan
        Write-Host "Named Bookable Resources with No Matching System User: $nNoMatchingSystemUsers" -ForegroundColor Cyan
        Write-Host "Named Bookable Resources Created:                      $nNamedBookableResourcesCreated" -ForegroundColor Green
        Write-Host "Named Bookable Resources Updated:                      $nNamedBookableResourcesUpdated" -ForegroundColor Yellow
        if ($Mode -eq 'NamedAndGeneric') {
            Write-Host "Unique Generic Resources Identified:                   $nProjectOnlineUniqueGenericResources" -ForegroundColor Cyan
            Write-Host "Generic Bookable Resources Created:                    $nGenericBookableResourcesCreated" -ForegroundColor Green
            Write-Host "Generic Bookable Resources Updated:                    $nGenericBookableResourcesUpdated" -ForegroundColor Yellow
        }
    }

    Write-Host "Import Resources completed" -ForegroundColor Green
}

function Get-BookableResources {
    <#
    .SYNOPSIS
        Retrieves all bookable resources from Dataverse.
    
    .DESCRIPTION
        Queries the sensei_bookableresources table in Dataverse to retrieve all named and generic resources.
        The function filters resources by type (955000001 for named resources and 955000000 for generic resources)
        and expands the user relationship to include domain names for correlation with system users.
        
        This function is used to check for existing resources in Dataverse before attempting to create new ones,
        helping to prevent duplicate resource creation during migration.
    
    .OUTPUTS
        System.Collections.Generic.List[PSCustomObject]
        Returns a collection of bookable resource records, each containing:
        - sensei_name: The resource name
        - sensei_user: Expanded user object with domain name (for named resources)
        - sensei_resourcetype: The resource type (955000001 = Named, 955000000 = Generic)
    
    .EXAMPLE
        $existingResources = Get-BookableResources
        
        Retrieves all bookable resources from Dataverse and stores them in the $existingResources variable.
    
    .NOTES
        Dependencies: Requires Get-AllRecords function from CommonFunctions.ps1
        Filters: Only retrieves resources with type 955000001 (Named) or 955000000 (Generic)
        Performance: Uses OData expand to minimize round-trip calls to Dataverse
    #>
    Write-Host '--Retrieving Bookable Resources from Dataverse--'

    $bookableResources = Get-AllRecords `
        -setName 'sensei_bookableresources' `
        -query '?$select=sensei_name&$expand=sensei_user($select=domainname)&$filter=sensei_resourcetype%20eq%20955000001%20or%20sensei_resourcetype%20eq%20955000000'
    
    return $bookableResources
}

function Get-SystemUsers {
    <#
    .SYNOPSIS
        Retrieves all system users from Dataverse.
    
    .DESCRIPTION
        Queries the systemusers table in Dataverse to retrieve all real user accounts (excluding application users)
        that have domain names. The function filters out application identities and users without domain names,
        returning only actual user accounts that can be matched with named resources during migration.
        
        This function is used to correlate Project Online resources with existing Dataverse system users by
        comparing domain names, enabling proper resource-to-user assignments during the migration process.
    
    .OUTPUTS
        System.Collections.Generic.List[PSCustomObject]
        Returns a collection of system user records, each containing:
        - fullname: The user's full name
        - domainname: The user's domain name (e.g., DOMAIN\username)
    
    .EXAMPLE
        $systemUsers = Get-SystemUsers
        
        Retrieves all system users from Dataverse and stores them in the $systemUsers variable.
    
    .NOTES
        Dependencies: Requires Get-AllRecords function from CommonFunctions.ps1
        Filters: Excludes application users (applicationid eq null) and users without domain names (domainname ne null)
        Use Case: Primarily used for matching named resources to existing Dataverse users during migration
    #>
    Write-Host '--Retrieving System Users from Dataverse--'

    $systemUsers = Get-AllRecords `
        -setName 'systemusers' `
        -query '?$select=fullname,domainname&$filter=applicationid eq null and domainname ne null'

    return $systemUsers
}
