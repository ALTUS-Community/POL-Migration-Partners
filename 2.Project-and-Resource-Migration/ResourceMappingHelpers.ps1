# ------------------------------
# CUSTOM FIELD, LOOKUP AND CHOICE MAPPING FUNCTIONS
# ------------------------------

function Get-ResourceCustomFieldTextValue {
    <#
    .SYNOPSIS
        Retrieves a single custom field text value from a Project Online resource.
    
    .DESCRIPTION
        Extracts the text value of a specified custom field from a Project Online resource object.
        The function safely navigates the resource's CustomFields collection, handling null values and missing
        properties at each level to prevent errors during data extraction.
        
        This is a core helper function used throughout the resource migration process to retrieve custom field data
        from Project Online resources for mapping to Dataverse bookable resource fields.
    
    .PARAMETER ProjectResource
        The Project Online resource object containing custom field data. 
    
    .PARAMETER CustomFieldName
        The name of the custom field to retrieve from the resource.
    
    .OUTPUTS
        System.String
        Returns the text value of the custom field, or $null if the field doesn't exist or has no value.
    
    .EXAMPLE
        $department = Get-ResourceCustomFieldTextValue -ProjectResource $resource -CustomFieldName 'Department'
        
        Retrieves the value of the 'Department' custom field from the resource.
    
    .NOTES
        Safety: Performs null checks at each level to prevent null reference exceptions
        Data Structure: Expects CustomFields to be a collection with CustomFieldValue.'#text' properties
        Use Case: Primarily used for single-value custom fields (text, numbers, dates)
        Difference from Projects: Uses CustomFieldValue instead of CFValue property
    #>
    param(
        [Parameter(Mandatory = $false)]
        $ProjectResource,

        [Parameter(Mandatory = $true)]
        [string]$CustomFieldName
    )

    if (-not $ProjectResource) { return $null }
    if (-not ($ProjectResource.PSObject.Properties.Name -contains 'CustomFields')) { return $null }
    if (-not $ProjectResource.CustomFields) { return $null }

    $cf = $ProjectResource.CustomFields | Where-Object { $_.CustomFieldName -eq $CustomFieldName } | Select-Object -First 1
    if (-not $cf) { return $null }
    if (-not ($cf.PSObject.Properties.Name -contains 'CustomFieldValue')) { return $null }
    if (-not $cf.CustomFieldValue) { return $null }

    return $cf.CustomFieldValue.'#text'
}

function Get-ResourceCustomFieldTextMultipleValues {
    <#
    .SYNOPSIS
        Retrieves multiple custom field text values from a Project Online resource.
    
    .DESCRIPTION
        Extracts all text values for a specified multi-value custom field from a Project Online resource object.
        The function handles custom fields that can have multiple selections or values, safely navigating the resource's
        CustomFields collection and handling both single objects and arrays of objects.
        
        This helper function is essential for mapping multi-select custom fields (like lookup tables with multiple selections)
        from Project Online resources to Dataverse multi-select choice fields or many-to-many relationships.
    
    .PARAMETER ProjectResource
        The Project Online resource object containing custom field data.
    
    .PARAMETER CustomFieldName
        The name of the multi-value custom field to retrieve from the resource.
    
    .OUTPUTS
        System.Array
        Returns an array of text values for the custom field, or $null if the field doesn't exist or has no values.
    
    .EXAMPLE
        $skills = Get-ResourceCustomFieldTextMultipleValues -ProjectResource $resource -CustomFieldName 'Skills'
        
        Retrieves all selected values from the 'Skills' multi-select custom field.
    
    .NOTES
        Safety: Performs null checks at each level to prevent null reference exceptions
        Data Structure: Handles both single object and array of objects for CustomFieldValue properties
        Use Case: Primarily used for multi-select custom fields (lookup tables, multi-choice fields)
        Return Type: Always returns an array, even for single values
        Difference from Projects: Uses CustomFieldValue instead of CFValue property
    #>
    param(
        [Parameter(Mandatory = $false)]
        $ProjectResource,

        [Parameter(Mandatory = $true)]
        [string]$CustomFieldName
    )

    if (-not $ProjectResource) { return $null }
    if (-not ($ProjectResource.PSObject.Properties.Name -contains 'CustomFields')) { return $null }
    if (-not $ProjectResource.CustomFields) { return $null }

    $cfValues = $ProjectResource.CustomFields | Where-Object { $_.CustomFieldName -eq $CustomFieldName }
    if (-not $cfValues -or $cfValues.Count -eq 0) { return $null }
    
    # Handle both single object and array of objects
    $firstItem = if ($cfValues -is [Array]) { $cfValues[0] } else { $cfValues }
    if (-not ($firstItem.PSObject.Properties.Name -contains 'CustomFieldValue')) { return $null }
    if (-not $cfValues.CustomFieldValue) { return $null }

    return @($cfValues.CustomFieldValue.'#text')
}

function Set-ResourceLookupField {
    <#
    .SYNOPSIS
        Maps a Project Online resource custom field to a Dataverse lookup field.
    
    .DESCRIPTION
        Retrieves a custom field value from a Project Online resource and maps it to a Dataverse lookup field
        by resolving the text value to a lookup table entry. The function automatically handles the OData binding syntax
        required for Dataverse lookup fields and adds the reference to the resource body hashtable.
        
        This is a higher-level helper function that combines custom field retrieval and lookup resolution into a single
        operation, simplifying the field mapping process during resource migration.
    
    .PARAMETER ResourceBody
        The hashtable representing the Dataverse bookable resource record body being constructed.
        The lookup reference will be added to this hashtable with the appropriate OData binding syntax.
    
    .PARAMETER ProjectResource
        The Project Online resource object containing the custom field data.
    
    .PARAMETER POLCustomFieldName
        The name of the custom field in Project Online to retrieve the value from.
    
    .PARAMETER LookupName
        The name of the lookup table in Dataverse to resolve the value against.
        Must match a lookup table name loaded via Get-LookupTableData.
    
    .PARAMETER DataverseFieldName
        The logical name of the Dataverse lookup field to populate (without the @odata.bind suffix).
    
    .EXAMPLE
        Set-ResourceLookupField -ResourceBody $resourceBody -ProjectResource $resource `
            -POLCustomFieldName 'Department' -LookupName 'Department' -DataverseFieldName 'sensei_department'
        
        Maps the 'Department' custom field from the resource to the sensei_department lookup field in Dataverse.
    
    .NOTES
        Dependencies: Requires Get-ResourceCustomFieldTextValue and Get-LookupValue functions
        OData Binding: Automatically appends '@odata.bind' suffix to the Dataverse field name
        Validation: Only adds the lookup reference if a valid match is found in the lookup table
        Safety: Performs null checks to prevent errors when values don't exist or can't be resolved
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ResourceBody,

        [Parameter(Mandatory = $true)]
        $ProjectResource,

        [Parameter(Mandatory = $true)]
        [string]$POLCustomFieldName,

        [Parameter(Mandatory = $true)]
        [string]$LookupName,

        [Parameter(Mandatory = $true)]
        [string]$DataverseFieldName
    )

    $textValue = Get-ResourceCustomFieldTextValue -ProjectResource $ProjectResource -CustomFieldName $POLCustomFieldName
    if ($textValue) {
        $lookupReference = Get-LookupValue -Value $textValue -LookupName $LookupName
        if ($lookupReference -and $null -ne $lookupReference -and $lookupReference -ne '') {
            $ResourceBody["$DataverseFieldName@odata.bind"] = $lookupReference
        }
    }
}

function Set-ResourceChoiceField {
    <#
    .SYNOPSIS
        Maps a Project Online resource custom field to a Dataverse choice (option set) field.
    
    .DESCRIPTION
        Retrieves a custom field value from a Project Online resource and maps it to a Dataverse choice field
        by resolving the text label to its corresponding choice value. The function supports both single-select and
        multi-select choice fields, automatically handling the comma-separated format required for multi-select choices.
        
        For single-select fields, the function retrieves one value and sets the choice field directly.
        For multi-select fields, it retrieves multiple values, resolves each to its choice value, and joins them
        with commas as required by Dataverse multi-select choice fields.
    
    .PARAMETER ResourceBody
        The hashtable representing the Dataverse bookable resource record body being constructed.
        The choice value(s) will be added to this hashtable.
    
    .PARAMETER ProjectResource
        The Project Online resource object containing the custom field data.
    
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
        Set-ResourceChoiceField -ResourceBody $resourceBody -ProjectResource $resource `
            -POLCustomFieldName 'Resource Type' -ChoiceName 'ResourceType' -DataverseFieldName 'sensei_resourcetype'
        
        Maps the 'Resource Type' custom field from the resource to a single-select choice field in Dataverse.
    
    .EXAMPLE
        Set-ResourceChoiceField -ResourceBody $resourceBody -ProjectResource $resource `
            -POLCustomFieldName 'Skills' -ChoiceName 'ResourceSkills' -DataverseFieldName 'sensei_skills' -IsMulti $true
        
        Maps the 'Skills' multi-select custom field from the resource to a multi-select choice field in Dataverse.
    
    .NOTES
        Dependencies: Requires Get-ResourceCustomFieldTextValue, Get-ResourceCustomFieldTextMultipleValues, and Get-ChoiceValue functions
        Multi-Select Format: Multi-select values are joined with commas (e.g., "100000000,100000001,100000002")
        Validation: Only adds choice values that can be successfully resolved against the choice field metadata
        Safety: Performs null checks to prevent errors when values don't exist or can't be resolved
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ResourceBody,

        [Parameter(Mandatory = $true)]
        $ProjectResource,

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
        $selectedValues = Get-ResourceCustomFieldTextMultipleValues -ProjectResource $ProjectResource -CustomFieldName $POLCustomFieldName
        if ($selectedValues) {
            $choiceValues = @()
            foreach ($val in $selectedValues) {
                $choiceValue = Get-ChoiceValue -Label $val -ChoiceName $ChoiceName
                if ($null -ne $choiceValue) {
                    $choiceValues += $choiceValue
                }
            }
            if ($choiceValues.Count -gt 0) {
                $ResourceBody[$DataverseFieldName] = $choiceValues -join ','
            }
        }
    } else {
        $textValue = Get-ResourceCustomFieldTextValue -ProjectResource $ProjectResource -CustomFieldName $POLCustomFieldName
        if ($textValue) {
            $choiceValue = Get-ChoiceValue -Label $textValue -ChoiceName $ChoiceName
            if ($null -ne $choiceValue) {
                $ResourceBody[$DataverseFieldName] = $choiceValue
            }
        }
    }
}

function Set-ResourceManyToManyRelationship {
    <#
    .SYNOPSIS
        Maps a Project Online resource multi-value custom field to a Dataverse many-to-many relationship.
    
    .DESCRIPTION
        Retrieves multiple custom field values from a Project Online resource and prepares them for
        creation as many-to-many relationships in Dataverse. The function resolves each text value to a related
        record ID using lookup tables and stores the relationship metadata in the resource body for deferred processing.
        
        Unlike direct field mappings, many-to-many relationships cannot be created during the initial resource creation.
        This function collects all relationship data into a special '_manyToManyRelationships' collection within the
        resource body hashtable, which is later extracted and processed by Split-ResourceManyToManyRelationships and
        Add-ResourceManyToManyRelationships after the resource record is created in Dataverse.
    
    .PARAMETER ResourceBody
        The hashtable representing the Dataverse bookable resource record body being constructed.
        Relationship metadata will be stored in a '_manyToManyRelationships' collection within this hashtable.
    
    .PARAMETER ProjectResource
        The Project Online resource object containing the custom field data.
    
    .PARAMETER POLCustomFieldName
        The name of the multi-value custom field in Project Online to retrieve values from.
    
    .PARAMETER LookupName
        The name of the lookup table as defined in Defaults.ps1.
        Must match a lookup table name loaded via Get-LookupTableData.
    
    .PARAMETER RelationshipName
        The logical name of the many-to-many relationship collection in Dataverse.
    
    .PARAMETER RelatedTableCollectionName
        The entity set name (collection name) of the related table in Dataverse.
    
    .PARAMETER RelatedTablePrimaryKey
        The primary key field name of the related table.
    
    .EXAMPLE
        Set-ResourceManyToManyRelationship -ResourceBody $ResourceBody -ProjectResource $ProjectResource`
            -POLCustomFieldName 'Multi Skills' -LookupName 'Skills' `
            -RelationshipName 'cr62c_sensei_bookableresource_cr62c_Skills_cr62c_Skills' `
            -RelatedTableCollectionName 'cr62c_skillses'`
            -RelatedTablePrimaryKey 'cr62c_skillsid'            
        Maps the 'Multi Skills' multi-select custom field to a many-to-many relationship with a custom skills table.
    
    .NOTES
        Dependencies: Requires Get-ResourceCustomFieldTextMultipleValues and Get-LookupValue functions
        Deferred Processing: Relationships are not created immediately; they're stored for later processing
        Workflow: Use with Split-ResourceManyToManyRelationships and Add-ResourceManyToManyRelationships
        Data Structure: Stores array of hashtables with CollectionName, SetName, PrimaryKey, and Id properties
        Safety: Performs null checks and only stores relationships for successfully resolved record IDs
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ResourceBody,

        [Parameter(Mandatory = $true)]
        $ProjectResource,

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

    $selectedValues = Get-ResourceCustomFieldTextMultipleValues -ProjectResource $ProjectResource -CustomFieldName $POLCustomFieldName

    if ($selectedValues) {
        Write-Host "Found $($selectedValues.Count) value(s) for field '$POLCustomFieldName'" -ForegroundColor Cyan
        foreach ($val in $selectedValues) {
            $relatedRecordId = Get-LookupValue -Value $val -LookupName $LookupName -ReturnIDOnly $true
            if ($null -ne $relatedRecordId) {
                # Store the relationship data to be processed after resource creation
                if (-not $ResourceBody.ContainsKey('_manyToManyRelationships')) {
                    $ResourceBody['_manyToManyRelationships'] = @()
                }
                $ResourceBody['_manyToManyRelationships'] += @{
                    CollectionName = $RelationshipName
                    SetName        = $RelatedTableCollectionName
                    PrimaryKey     = $RelatedTablePrimaryKey
                    Id             = $relatedRecordId
                }
            }
        }
    }
}

function Split-ResourceManyToManyRelationships {
    <#
    .SYNOPSIS
        Extracts and removes many-to-many relationship metadata from a resource body hashtable.
    
    .DESCRIPTION
        Separates the many-to-many relationship metadata from a Dataverse bookable resource record body hashtable.
        The function retrieves the '_manyToManyRelationships' collection (if present) and removes it from the body,
        allowing the body to be used for record creation/update operations while preserving the relationship data
        for separate processing.
        
        This separation is necessary because many-to-many relationships cannot be created during the initial
        record creation in Dataverse. They must be processed separately using the Web API after the parent
        record has been created and assigned an ID.
    
    .PARAMETER Body
        The hashtable representing the Dataverse bookable resource record body.
        If it contains a '_manyToManyRelationships' key, that collection will be extracted and removed.
    
    .OUTPUTS
        System.Array
        Returns the array of many-to-many relationship metadata objects, or $null if no relationships were stored.
        Each relationship object contains: CollectionName, SetName, PrimaryKey, and Id properties.
    
    .EXAMPLE
        $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
        $resourceId = New-Record -setName 'sensei_bookableresources' -body $resourceBody
        Add-ResourceManyToManyRelationships -ResourceId $resourceId -ManyToManyRelationships $manyToManyRelationships
        
        Extracts relationship metadata, creates the resource record, then processes the relationships.
    
    .NOTES
        Dependencies: Works in conjunction with Set-ResourceManyToManyRelationship and Add-ResourceManyToManyRelationships
        Side Effect: Modifies the input Body hashtable by removing the '_manyToManyRelationships' key
        Workflow: Called after building resource body and before New-Record or Update-Record
        Return Value: Safe to pass $null to Add-ResourceManyToManyRelationships if no relationships exist
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

function Add-ResourceManyToManyRelationships {
    <#
    .SYNOPSIS
        Creates many-to-many relationships between a bookable resource and related records in Dataverse.
    
    .DESCRIPTION
        Processes the many-to-many relationship metadata extracted from a resource body and creates the actual
        relationships in Dataverse using the Web API. The function groups relationships by collection to minimize
        API calls, checks for existing relationships to prevent duplicates, and handles errors gracefully.
        
        This function must be called after the bookable resource record has been created in Dataverse and has a valid ID.
        It works in conjunction with Set-ResourceManyToManyRelationship (which collects relationship metadata)
        and Split-ResourceManyToManyRelationships (which extracts the metadata from the resource body).
        
        The function optimizes performance by:
        - Grouping relationships by collection name to batch operations
        - Querying existing relationships once per collection to avoid duplicate creation attempts
        - Using the primary key field to efficiently check for existing relationships
    
    .PARAMETER ResourceId
        The GUID of the bookable resource record in Dataverse that the relationships will be associated with.
        This must be a valid, existing resource record ID.
    
    .PARAMETER ManyToManyRelationships
        The array of relationship metadata objects returned from Split-ResourceManyToManyRelationships.
        Each object should contain: CollectionName, SetName, PrimaryKey, and Id properties.
        Can be $null if no relationships need to be created.
    
    .EXAMPLE
        $manyToManyRelationships = Split-ResourceManyToManyRelationships -Body $resourceBody
        $resourceId = New-Record -setName 'sensei_bookableresources' -body $resourceBody
        Add-ResourceManyToManyRelationships -ResourceId $resourceId -ManyToManyRelationships $manyToManyRelationships
        
        Standard workflow: Extract relationships, create resource record, then add the relationships.
    
    .NOTES
        Dependencies: Requires Get-Records and Add-ToCollection functions from CommonFunctions.ps1
        Performance: Groups relationships by collection to minimize queries and optimize API usage
        Error Handling: Continues processing remaining relationships even if individual relationships fail
        Duplicate Prevention: Checks existing relationships before attempting to create new ones
        Workflow: Called after New-Record or Update-Record has created/updated the parent resource record
    #>
    param(
        [Parameter(Mandatory = $true)]
        [guid]$ResourceId,

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
                $existingRelationships = (Get-Records -setName 'sensei_bookableresources' -query "($ResourceId)/$collectionName").value
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
                            -targetSetName 'sensei_bookableresources' `
                            -targetId $ResourceId `
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
