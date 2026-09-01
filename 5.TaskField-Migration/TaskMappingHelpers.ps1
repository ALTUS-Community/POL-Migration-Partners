# ------------------------------
# CUSTOM FIELD, LOOKUP AND CHOICE MAPPING FUNCTIONS
# ------------------------------

function Get-ReportingTaskCustomFieldTextValue {
    <#
    .SYNOPSIS
        Retrieves a single custom field text value from a Project Online reporting task.
    
    .DESCRIPTION
        Extracts the text value of a specified custom field from a Project Online reporting task object.
        The function safely navigates the task's CustomFields collection, handling null values and missing
        properties at each level to prevent errors during data extraction.
        
        This is a core helper function used throughout the migration process to retrieve custom field data
        from Project Online for mapping to Dataverse fields.
    
    .PARAMETER ReportingTask
        The Project Online reporting task object containing custom field data. Can be null.
    
    .PARAMETER CustomFieldName
        The name of the custom field to retrieve from the reporting task.
    
    .OUTPUTS
        System.String
        Returns the text value of the custom field, or $null if the field doesn't exist or has no value.
    
    .EXAMPLE
        $taskType = Get-ReportingCustomFieldTextValue -ReportingTask $reportingTask -CustomFieldName 'Task Type'
        
        Retrieves the value of the 'Task Type' custom field from the reporting task.
    
    .NOTES
        Safety: Performs null checks at each level to prevent null reference exceptions
        Data Structure: Expects CustomFields to be a collection with CFValue.'#text' properties
        Use Case: Primarily used for single-value custom fields (text, numbers, dates)
    #>
    param(
        [Parameter(Mandatory = $false)]
        $ReportingTask,

        [Parameter(Mandatory = $true)]
        [string]$CustomFieldName
    )

    if (-not $ReportingTask) { return $null }
    if (-not ($ReportingTask.PSObject.Properties.Name -contains 'CustomFields')) { return $null }
    if (-not $ReportingTask.CustomFields) { return $null }

    $cf = $ReportingTask.CustomFields | Where-Object { $_.CustomFieldName -eq $CustomFieldName } | Select-Object -First 1
    if (-not $cf) { return $null }
    if (-not ($cf.PSObject.Properties.Name -contains 'CustomFieldValue')) { return $null }
    if (-not $cf.CustomFieldValue) { return $null }

    return $cf.CustomFieldValue.'#text'
}

function Get-ReportingTaskCustomFieldTextMultipleValues {
    <#
    .SYNOPSIS
        Retrieves multiple custom field text values from a Project Online reporting task.
    
    .DESCRIPTION
        Extracts all text values for a specified multi-value custom field from a Project Online reporting task object.
        The function handles custom fields that can have multiple selections or values, safely navigating the task's
        CustomFields collection and handling both single objects and arrays of objects.
        
        This helper function is essential for mapping multi-select custom fields (like lookup tables with multiple selections)
        from Project Online to Dataverse multi-select choice fields or many-to-many relationships.
    
    .PARAMETER ReportingTask
        The Project Online reporting task object containing custom field data. 
    
    .PARAMETER CustomFieldName
        The name of the multi-value custom field to retrieve from the reporting task.
    
    .OUTPUTS
        System.Array
        Returns an array of text values for the custom field, or $null if the field doesn't exist or has no values.
    
    .EXAMPLE
        $categories = Get-ReportingCustomFieldTextMultipleValues -ReportingTask $reportingTask -CustomFieldName 'Task Categories'
        
        Retrieves all selected values from the 'Task Categories' multi-select custom field.
    
    .NOTES
        Safety: Performs null checks at each level to prevent null reference exceptions
        Data Structure: Handles both single object and array of objects for CFValue properties
        Use Case: Primarily used for multi-select custom fields (lookup tables, multi-choice fields)
        Return Type: Always returns an array, even for single values
    #>
    param(
        [Parameter(Mandatory = $false)]
        $ReportingTask,

        [Parameter(Mandatory = $true)]
        [string]$CustomFieldName
    )

    if (-not $ReportingTask) { return $null }
    if (-not ($ReportingTask.PSObject.Properties.Name -contains 'CustomFields')) { return $null }
    if (-not $ReportingTask.CustomFields) { return $null }

    $cfValues = $ReportingTask.CustomFields | Where-Object { $_.CustomFieldName -eq $CustomFieldName }
    if (-not $cfValues -or $cfValues.Count -eq 0) { return $null }
    
    # Handle both single object and array of objects
    $firstItem = if ($cfValues -is [Array]) { $cfValues[0] } else { $cfValues }
    if (-not ($firstItem.PSObject.Properties.Name -contains 'CustomFieldValue')) { return $null }
    if (-not $cfValues.CustomFieldValue) { return $null }

    return @($cfValues.CustomFieldValue.'#text')
}

function Set-TaskLookupField {
    <#
    .SYNOPSIS
        Maps a Project Online custom field to a Dataverse lookup field.
    
    .DESCRIPTION
        Retrieves a custom field value from a Project Online reporting task and maps it to a Dataverse lookup field
        by resolving the text value to a lookup table entry. The function automatically handles the OData binding syntax
        required for Dataverse lookup fields and adds the reference to the task body hashtable.
        
        This is a higher-level helper function that combines custom field retrieval and lookup resolution into a single
        operation, simplifying the field mapping process during task migration.
    
    .PARAMETER TaskBody
        The hashtable representing the Dataverse task record body being constructed.
        The lookup reference will be added to this hashtable with the appropriate OData binding syntax.
    
    .PARAMETER ReportingTask
        The Project Online reporting task object containing the custom field data.
    
    .PARAMETER POLCustomFieldName
        The name of the custom field in Project Online to retrieve the value from.
    
    .PARAMETER LookupName
        The name of the lookup table as defined in Defaults.ps1.
        Must match a lookup table name loaded via Get-LookupTableData.
    
    .PARAMETER DataverseFieldName
        The logical name of the Dataverse lookup field to populate (without the @odata.bind suffix).
    
    .EXAMPLE
        Set-TaskLookupField -TaskBody $taskBody -ReportingTask $reportingTask `
            -POLCustomFieldName 'Department' -LookupName 'Department' -DataverseFieldName 'sensei_department'
        
        Maps the 'Department' custom field from Project Online to the sensei_department lookup field in Dataverse.
    
    .NOTES
        Dependencies: Requires Get-ReportingCustomFieldTextValue and Get-LookupValue functions
        OData Binding: Automatically appends '@odata.bind' suffix to the Dataverse field name
        Validation: Only adds the lookup reference if a valid match is found in the lookup table
        Safety: Performs null checks to prevent errors when values don't exist or can't be resolved
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TaskBody,

        [Parameter(Mandatory = $true)]
        $ReportingTask,

        [Parameter(Mandatory = $true)]
        [string]$POLCustomFieldName,

        [Parameter(Mandatory = $true)]
        [string]$LookupName,

        [Parameter(Mandatory = $true)]
        [string]$DataverseFieldName
    )

    $textValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName $POLCustomFieldName
    if ($textValue) {
        $lookupReference = Get-LookupValue -Value $textValue -LookupName $LookupName
        if ($lookupReference -and $null -ne $lookupReference -and $lookupReference -ne '') {
            $TaskBody["$DataverseFieldName@odata.bind"] = $lookupReference
        }
    }
}

function Set-TaskChoiceField {
    <#
    .SYNOPSIS
        Maps a Project Online custom field to a Dataverse choice (option set) field.
    
    .DESCRIPTION
        Retrieves a custom field value from a Project Online reporting task and maps it to a Dataverse choice field
        by resolving the text label to its corresponding choice value. The function supports both single-select and
        multi-select choice fields, automatically handling the comma-separated format required for multi-select choices.
        
        For single-select fields, the function retrieves one value and sets the choice field directly.
        For multi-select fields, it retrieves multiple values, resolves each to its choice value, and joins them
        with commas as required by Dataverse multi-select choice fields.
    
    .PARAMETER TaskBody
        The hashtable representing the Dataverse task record body being constructed.
        The choice value(s) will be added to this hashtable.
    
    .PARAMETER ReportingTask
        The Project Online reporting task object containing the custom field data.
    
    .PARAMETER POLCustomFieldName
        The name of the custom field in Project Online to retrieve the value from.
    
    .PARAMETER ChoiceName
        The name of the choice field (option set) in Dataverse to resolve the label against.
        Must match a choice field name loaded via Get-ChoiceFieldData.
    
    .PARAMETER DataverseFieldName
        The logical name of the Dataverse choice field to populate.
    
    .PARAMETER IsMulti
        Indicates whether the choice field supports multiple selections.
        Default is $false (single-select). Set to $true for multi-select choice fields.
    
    .EXAMPLE
        Set-TaskChoiceField -TaskBody $taskBody -ReportingTask $reportingTask `
            -POLCustomFieldName 'Status' -ChoiceName 'TaskStatus' -DataverseFieldName 'sensei_status'
        
        Maps the 'Status' custom field from Project Online to a single-select choice field in Dataverse.
    
    .EXAMPLE
        Set-TaskChoiceField -TaskBody $taskBody -ReportingTask $reportingTask `
            -POLCustomFieldName 'Tags' -ChoiceName 'TaskTags' -DataverseFieldName 'sensei_tags' -IsMulti $true
        
        Maps the 'Tags' multi-select custom field from Project Online to a multi-select choice field in Dataverse.
    
    .NOTES
        Dependencies: Requires Get-ReportingTaskCustomFieldTextValue, Get-ReportingTaskCustomFieldTextMultipleValues, and Get-ChoiceValue functions
        Multi-Select Format: Multi-select values are joined with commas (e.g., "100000000,100000001,100000002")
        Validation: Only adds choice values that can be successfully resolved against the choice field metadata
        Safety: Performs null checks to prevent errors when values don't exist or can't be resolved
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TaskBody,

        [Parameter(Mandatory = $true)]
        $ReportingTask,

        [Parameter(Mandatory = $true)]
        [string]$POLCustomFieldName,

        [Parameter(Mandatory = $true)]
        [string]$ChoiceName,

        [Parameter(Mandatory = $true)]
        [string]$DataverseFieldName,

        [Parameter(Mandatory = $false)]
        [bool]$IsMulti = $false
    )

    if ($IsMulti) {
        $selectedValues = Get-ReportingTaskCustomFieldTextMultipleValues -ReportingTask $ReportingTask -CustomFieldName $POLCustomFieldName
        if ($selectedValues) {
            $choiceValues = @()
            foreach ($val in $selectedValues) {
                $choiceValue = Get-ChoiceValue -Label $val -ChoiceName $ChoiceName
                if ($null -ne $choiceValue) {
                    $choiceValues += $choiceValue
                }
            }
            if ($choiceValues.Count -gt 0) {
                $TaskBody[$DataverseFieldName] = $choiceValues -join ','
            }
        }
    } else {
        $textValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName $POLCustomFieldName
        if ($textValue) {
            $choiceValue = Get-ChoiceValue -Label $textValue -ChoiceName $ChoiceName
            if ($null -ne $choiceValue) {
                $TaskBody[$DataverseFieldName] = $choiceValue
            }
        }
    }
}

function Set-TaskManyToManyRelationship {
    <#
    .SYNOPSIS
        Maps a Project Online multi-value custom field to a Dataverse many-to-many relationship.
    
    .DESCRIPTION
        Retrieves multiple custom field values from a Project Online reporting task and prepares them for
        creation as many-to-many relationships in Dataverse. The function resolves each text value to a related
        record ID using lookup tables and stores the relationship metadata in the task body for deferred processing.
        
        Unlike direct field mappings, many-to-many relationships cannot be created during the initial task creation.
        This function collects all relationship data into a special '_manyToManyRelationships' collection within the
        task body hashtable, which is later extracted and processed by Split-ManyToManyRelationships and
        Add-TaskManyToManyRelationships after the task record is created in Dataverse.
    
    .PARAMETER TaskBody
        The hashtable representing the Dataverse task record body being constructed.
        Relationship metadata will be stored in a '_manyToManyRelationships' collection within this hashtable.
    
    .PARAMETER ReportingTask
        The Project Online reporting task object containing the custom field data.
    
    .PARAMETER POLCustomFieldName
        The name of the multi-value custom field in Project Online to retrieve values from.
    
    .PARAMETER LookupName
        The name of the lookup table in Dataverse to resolve the text values against.
        Must match a lookup table name loaded via Get-LookupTableData.
    
    .PARAMETER RelationshipName
        The logical name of the many-to-many relationship collection in Dataverse (e.g., 'sensei_project_sensei_department').
    
    .PARAMETER RelatedTableCollectionName
        The entity set name (collection name) of the related table in Dataverse (e.g., 'sensei_departments').
    
    .PARAMETER RelatedTablePrimaryKey
        The primary key field name of the related table (e.g., 'sensei_departmentid').
    
    .EXAMPLE
        Set-TaskManyToManyRelationship -TaskBody $Task -ReportingTask $ReportingTask `
            -POLCustomFieldName 'Multi Department' -LookupName 'Department' `
            -RelationshipName 'cr62c_sensei_task_cr62c_Department_cr62c_Department' `
            -RelatedTableCollectionName 'cr62c_departments' `
            -RelatedTablePrimaryKey 'cr62c_departmentid'
        Maps the 'Multi Department' multi-select custom field to a many-to-many relationship with a custom departments table.
    
    .NOTES
        Dependencies: Requires Get-ReportingTaskCustomFieldTextMultipleValues and Get-LookupValue functions
        Deferred Processing: Relationships are not created immediately; they're stored for later processing
        Workflow: Use with Split-ManyToManyRelationships and Add-TaskManyToManyRelationships
        Data Structure: Stores array of hashtables with CollectionName, SetName, PrimaryKey, and Id properties
        Safety: Performs null checks and only stores relationships for successfully resolved record IDs
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TaskBody,

        [Parameter(Mandatory = $true)]
        $ReportingTask,

        [Parameter(Mandatory = $true)]
        [string]$POLCustomFieldName,

        [Parameter(Mandatory = $true)]
        [string]$LookupName,

        [Parameter(Mandatory = $true)]
        [string]$RelationshipName,

        [Parameter(Mandatory = $true)]
        [string]$RelatedTableCollectionName,

        [Parameter(Mandatory = $true)]
        [string]$RelatedTablePrimaryKey
    )

    $selectedValues = Get-ReportingTaskCustomFieldTextMultipleValues -ReportingTask $ReportingTask -CustomFieldName $POLCustomFieldName

    if ($selectedValues) {
        # Write-Host "Found $($selectedValues.Count) value(s) for field '$POLCustomFieldName'" -ForegroundColor Cyan
        foreach ($val in $selectedValues) {
            $relatedRecordId = Get-LookupValue -Value $val -LookupName $LookupName -ReturnIDOnly $true
            if ($null -ne $relatedRecordId) {
                # Store the relationship data to be processed after task creation
                if (-not $TaskBody.ContainsKey('_manyToManyRelationships')) {
                    $TaskBody['_manyToManyRelationships'] = @()
                }
                $TaskBody['_manyToManyRelationships'] += @{
                    CollectionName = $RelationshipName
                    SetName        = $RelatedTableCollectionName
                    PrimaryKey     = $RelatedTablePrimaryKey
                    Id             = $relatedRecordId
                }
            }
        }
    } else {
        # Write-Host "No values found for field '$POLCustomFieldName'" -ForegroundColor Yellow
    }
}

function Split-TaskManyToManyRelationships {
    <#
    .SYNOPSIS
        Extracts and removes many-to-many relationship metadata from a task body hashtable.
    
    .DESCRIPTION
        Separates the many-to-many relationship metadata from a Dataverse task record body hashtable.
        The function retrieves the '_manyToManyRelationships' collection (if present) and removes it from the body,
        allowing the body to be used for record creation/update operations while preserving the relationship data
        for separate processing.
        
        This separation is necessary because many-to-many relationships cannot be created during the initial
        record creation in Dataverse. They must be processed separately using the Web API after the parent
        record has been created and assigned an ID.
    
    .PARAMETER Body
        The hashtable representing the Dataverse task record body.
        If it contains a '_manyToManyRelationships' key, that collection will be extracted and removed.
    
    .OUTPUTS
        System.Array
        Returns the array of many-to-many relationship metadata objects, or $null if no relationships were stored.
        Each relationship object contains: CollectionName, SetName, PrimaryKey, and Id properties.
    
    .EXAMPLE
        $manyToManyRelationships = Split-TaskManyToManyRelationships -Body $taskBody
        $taskId = New-Record -setName 'sensei_tasks' -body $taskBody
        Add-TaskManyToManyRelationships -TaskId $taskId -ManyToManyRelationships $manyToManyRelationships
        
        Extracts relationship metadata, creates the task record, then processes the relationships.
    
    .NOTES
        Dependencies: Works in conjunction with Set-TaskManyToManyRelationship and Add-TaskManyToManyRelationships
        Side Effect: Modifies the input Body hashtable by removing the '_manyToManyRelationships' key
        Workflow: Called after Get-TaskDataverseBody and before New-Record or Update-Record
        Return Value: Safe to pass $null to Add-TaskManyToManyRelationships if no relationships exist
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Body
    )

    $manyToManyRelationships = $null
    if ($Body.ContainsKey('_manyToManyRelationships')) {
        $manyToManyRelationships = $Body['_manyToManyRelationships']
        $Body.Remove('_manyToManyRelationships')
    }
    return $manyToManyRelationships
}

function Add-TaskManyToManyRelationships {
    <#
    .SYNOPSIS
        Creates many-to-many relationships between a task and related records in Dataverse.
    
    .DESCRIPTION
        Processes the many-to-many relationship metadata extracted from a task body and creates the actual
        relationships in Dataverse using the Web API. The function groups relationships by collection to minimize
        API calls, checks for existing relationships to prevent duplicates, and handles errors gracefully.
        
        This function must be called after the task record has been created in Dataverse and has a valid ID.
        It works in conjunction with Set-TaskManyToManyRelationship (which collects relationship metadata)
        and Split-TaskManyToManyRelationships (which extracts the metadata from the task body).
        
        The function optimizes performance by:
        - Grouping relationships by collection name to batch operations
        - Querying existing relationships once per collection to avoid duplicate creation attempts
        - Using the primary key field to efficiently check for existing relationships
    
    .PARAMETER TaskId
        The GUID of the task record in Dataverse that the relationships will be associated with.
        This must be a valid, existing task record ID.
    
    .PARAMETER ManyToManyRelationships
        The array of relationship metadata objects returned from Split-TaskManyToManyRelationships.
        Each object should contain: CollectionName, SetName, PrimaryKey, and Id properties.
        Can be $null if no relationships need to be created.
    
    .EXAMPLE
        $manyToManyRelationships = Split-TaskManyToManyRelationships -Body $taskBody
        $taskId = New-Record -setName 'sensei_tasks' -body $taskBody
        Add-TaskManyToManyRelationships -TaskId $taskId -ManyToManyRelationships $manyToManyRelationships
        
        Standard workflow: Extract relationships, create task record, then add the relationships.
    
    .NOTES
        Dependencies: Requires Get-Records and Add-ToCollection functions from CommonFunctions.ps1
        Performance: Groups relationships by collection to minimize queries and optimize API usage
        Error Handling: Continues processing remaining relationships even if individual relationships fail
        Duplicate Prevention: Checks existing relationships before attempting to create new ones
        Workflow: Called after New-Record or Update-Record has created/updated the parent task record
    #>
    param(
        [Parameter(Mandatory = $true)]
        [guid]$TaskId,

        [Parameter(Mandatory = $false)]
        $ManyToManyRelationships
    )

    if ($null -ne $ManyToManyRelationships) {
        # Group relationships by collection name to minimize queries
        $relationshipsByCollection = $ManyToManyRelationships | Group-Object -Property CollectionName
        
        foreach ($group in $relationshipsByCollection) {
            $collectionName = $group.Name
            
            # Get all existing relationships for this collection in one query
            try {
                $existingRelationships = (Get-Records -setName 'sensei_tasks' -query "($TaskId)/$collectionName").value
                $existingIds = @{}
                
                # Get the primary key field name from the first relationship in this group
                $primaryKeyField = $group.Group[0].PrimaryKey
                
                if ($existingRelationships) {
                    foreach ($rel in $existingRelationships) {
                        if ($rel.PSObject.Properties.Name -contains $primaryKeyField) {
                            $existingIds[$rel.$primaryKeyField] = $true
                        }
                    }
                }
            } catch {
                Write-Host "Warning: Could not retrieve existing relationships for $collectionName. Will attempt to add all." -ForegroundColor Yellow
                $existingIds = @{}
            }
            
            # Process each relationship in this collection
            foreach ($relationship in $group.Group) {
                if (-not $existingIds.ContainsKey($relationship.Id)) {
                    #Write-Host "Adding many-to-many relationship to $($relationship.SetName) with Id $($relationship.Id)..." -ForegroundColor Cyan
                    try {
                        Add-ToCollection `
                            -targetSetName 'sensei_tasks' `
                            -targetId $TaskId `
                            -collectionName $relationship.CollectionName `
                            -setName $relationship.SetName `
                            -id $relationship.Id
                    } catch {
                        Write-Host "Warning: Failed to add relationship: $_" -ForegroundColor Yellow
                    }
                } else {
                    #Write-Host "Many-to-many relationship to $($relationship.SetName) with Id $($relationship.Id) already exists, skipping..." -ForegroundColor Gray
                }
            }
        }
    }
}

function Remove-BlocklistedTaskFields {
    <#
    .SYNOPSIS
        Removes blocklisted fields from a task body hashtable before an Update operation.

    .DESCRIPTION
        Certain Dataverse task fields must not be updated via the migration process as doing so
        would cause unintended side effects on system functionality. This function strips any
        such fields from the task body hashtable before it is passed to Update-Record.

        Add fields to the $blocklist array below to prevent them from being written.

    .PARAMETER TaskBody
        The hashtable representing the Dataverse task record to be updated. Modified in place.

    .EXAMPLE
        Remove-BlocklistedTaskFields -TaskBody $updateBody
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$TaskBody
    )

    $blocklist = @(
        'sensei_duration',
        'sensei_effort',
        'sensei_effortcompleted',
        'sensei_effortremaining',
        'sensei_effortremainingoverride',
        'sensei_externalparenttaskid',
        'sensei_externalprojectid',
        'sensei_externaltaskid',
        'sensei_fieldlock',
        'sensei_formupdated',
        'sensei_hash',
        'sensei_index',
        'sensei_isexternal',
        'sensei_json_properties',
        'sensei_lastsyncdate',
        'sensei_level',
        'sensei_locked',
        'sensei_milestone',
        'sensei_name',
        'sensei_notes',
        'sensei_order',
        'sensei_originofchange',
        'sensei_parenttask',
        'sensei_percentcomplete',
        'sensei_percentcompletedecimal',
        'sensei_project',
        'sensei_scheduletype',
        'sensei_summary',
        'sensei_taskfinish',
        'sensei_taskfinishdatetime',
        'sensei_tasknumber',
        'sensei_taskstart',
        'sensei_taskstartdatetime',
        'sensei_tasktype'
    )

    foreach ($field in $blocklist) {
        if ($TaskBody.ContainsKey($field)) {
            Write-Host "    Removing blocklisted field '$field' from task update body." -ForegroundColor DarkYellow
            $TaskBody.Remove($field)
        }
    }
}
