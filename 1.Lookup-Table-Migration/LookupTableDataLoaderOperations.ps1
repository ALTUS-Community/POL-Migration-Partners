# ==============================================================================
# LOOKUP TABLE DATA IMPORT
# ==============================================================================

<#
.SYNOPSIS
Imports lookup table data from Project Online to Dataverse.

.DESCRIPTION
The Import-LookupTableData function takes lookup table data exported from Project Online and imports it into 
specified Dataverse tables. It supports both flat and hierarchical mapping structures.

.PARAMETER LookupData
The lookup table data imported from CSV. Should contain columns: InternalName, FullValue, Description, SortIndex, Id.

.PARAMETER MapAsHierarchy
Boolean value indicating whether to map the data as a hierarchy ($true) or as a flat structure ($false).

.PARAMETER DataverseTable
The Dataverse table configuration object containing table and field mappings.

.EXAMPLE
Import-LookupTableData -LookupData $lookupData -MapAsHierarchy $false -DataverseTable $tableConfig

This example imports flat lookup table data to a single Dataverse table.

.EXAMPLE
Import-LookupTableData -LookupData $lookupData -MapAsHierarchy $true -DataverseTable $tableConfig

This example imports hierarchical lookup table data with parent-child relationships.
#>

function Import-LookupTableData {
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$LookupData,
        
        [Parameter(Mandatory = $false)]
        [bool]$MapAsHierarchy,

        [Parameter(Mandatory = $false)]
        [object[]]$DataverseTable
    )
    
    # Lookup table mapping
    Write-Host "  Migration mode: $(if ($MapAsHierarchy) { 'Self Referential Hierarchy Table' } else { 'Flat Table' })" -ForegroundColor Gray
    Write-Host "  Target table: $($DataverseTable.TableCollectionName)" -ForegroundColor Gray
    Write-Host "  Total entries: $($LookupData.Count)" -ForegroundColor Gray

    $tableCollectionName = $DataverseTable.TableCollectionName
    $nameField = $DataverseTable.NameField
    $descriptionField = $DataverseTable.DescriptionField

    if ($MapAsHierarchy) {
        Write-Host "`nProcessing self-referential hierarchical mapping..." -ForegroundColor Cyan
        # Self-referential hierarchical mapping: all data goes to a single table with parent-child relationships
        
        $prevLevelLookup = $DataverseTable.PrevLevelLookupField
        
        Write-Host "  Target table: $tableCollectionName" -ForegroundColor Gray
        if ($prevLevelLookup) {
            Write-Host "  Parent lookup field: $prevLevelLookup" -ForegroundColor Gray
        }
        
        # Load code mask to determine separators
        $codeMaskFile = $sourceFile -replace '^LookupTable_', 'CodeMask_'
        $codeMaskPath = Join-Path $PSScriptRoot "LookupTables" $codeMaskFile
        $separators = @()
        
        if (Test-Path $codeMaskPath) {
            $codeMask = Import-Csv -Path $codeMaskPath -Encoding UTF8
            $separators = $codeMask | ForEach-Object { $_.Separator }
            Write-Host "  Code mask loaded: $($separators.Count) levels with separators: $($separators -join ', ')" -ForegroundColor Gray
        } else {
            Write-Host "  No code mask found, using default separators (. and \)" -ForegroundColor Yellow
        }
        
        # Group data by hierarchy level (determined by counting separators in FullValue)
        $hierarchyLevels = @{}
        foreach ($entry in $LookupData) {
            $level = 0
            if ($entry.FullValue -and $separators.Count -gt 0) {
                # Count hierarchy separators based on code mask
                $escapedSeparators = $separators | ForEach-Object { [regex]::Escape($_) }
                $pattern = '[' + ($escapedSeparators -join '') + ']'
                $level = ([regex]::Matches($entry.FullValue, $pattern)).Count
            } elseif ($entry.FullValue) {
                # Fallback: Count default separators (. or \)
                $level = ([regex]::Matches($entry.FullValue, '[.\\/]')).Count
            }
            
            if (-not $hierarchyLevels.ContainsKey($level)) {
                $hierarchyLevels[$level] = @()
            }
            $hierarchyLevels[$level] += $entry
        }
        
        # Track created records for parent-child lookups
        $createdRecords = @{}
        
        # Process each hierarchy level in order (top-level first)
        $levelKeys = $hierarchyLevels.Keys | Sort-Object
        foreach ($level in $levelKeys) {
            $entries = $hierarchyLevels[$level]
            
            Write-Host "`n  Processing level $level ($($entries.Count) entries)" -ForegroundColor Yellow
            
            foreach ($entry in $entries) {
                try {
                    # For self-referential hierarchies, use the full path as the name
                    $fullPathName = $entry.FullValue
                    
                    # Extract parent value and last separator index for parent lookup
                    $lastSeparatorIndex = $entry.FullValue.LastIndexOfAny(@('.', '\', '/'))
                    
                    # Create the record in Dataverse
                    $record = @{
                        $nameField        = $fullPathName
                        $descriptionField = $entry.Description
                    }
                    
                    # If this is not the top level and we have a parent lookup field, find and set the parent reference
                    if ($prevLevelLookup -and $level -gt 0) {
                        # Extract parent value (everything before the last separator)
                        if ($lastSeparatorIndex -gt 0) {
                            $parentValue = $entry.FullValue.Substring(0, $lastSeparatorIndex)
                            
                            # Find the parent record ID from previously created records
                            if ($createdRecords.ContainsKey($parentValue)) {
                                $parentRecordId = $createdRecords[$parentValue]
                                
                                # Add lookup reference to parent using OData bind syntax
                                $lookupBindKey = "$prevLevelLookup@odata.bind"
                                $lookupBindValue = "/$tableCollectionName($parentRecordId)"
                                $record[$lookupBindKey] = $lookupBindValue
                                
                                Write-Host "    Linking '$fullPathName' to parent '$parentValue' (ID: $parentRecordId)" -ForegroundColor Gray
                            } else {
                                Write-Host "    Warning: Parent record not found for '$parentValue'" -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    $newRecordId = New-Record -setName $tableCollectionName -body $record
                    
                    # Store the created record for potential child lookups
                    $createdRecords[$entry.FullValue] = $newRecordId
                    
                    Write-Host "    ✓ Created: $($entry.FullValue)" -ForegroundColor Green
                } catch {
                    Write-Host "    ✗ Failed to create '$($entry.FullValue)': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
    } else {
        # Flat mapping: all data goes to a single table
        Write-Host "`nProcessing flat mapping to table '$($tableCollectionName)'..." -ForegroundColor Cyan
        
        $successCount = 0
        $failureCount = 0
        
        foreach ($entry in $LookupData) {
            try {
                # Create the record in Dataverse
                $record = @{
                    $nameField        = $entry.FullValue
                    $descriptionField = $entry.Description
                }
                
                $newRecordId = New-Record -setName $tableCollectionName -body $record
                Write-Host "  ✓ Created: $($entry.FullValue)" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Host "  ✗ Failed to create '$($entry.FullValue)': $($_.Exception.Message)" -ForegroundColor Red
                $failureCount++
            }
        }
        
        Write-Host "`nFlat mapping completed: $successCount successful, $failureCount failed" -ForegroundColor $(if ($failureCount -eq 0) { 'Green' } else { 'Yellow' })
    }
}

# ==============================================================================
# ITEM CONVERSION
# ==============================================================================

<#
.SYNOPSIS
Converts a lookup table entry into a Dataverse entity record format.

.DESCRIPTION
Transforms a lookup table item (from CSV) into a hashtable structure suitable for creating 
or updating a Dataverse record. Handles both flat lookups and hierarchical self-referential 
lookups by mapping fields and establishing parent-child relationships based on code mask separators.

For hierarchical lookups, the function parses the FullValue to determine the parent entry 
and creates the appropriate lookup reference.

.PARAMETER Item
The lookup table item object containing FullValue, Description, and other properties from the CSV.

.PARAMETER LookupTableMapping
The mapping configuration from LookupTablesToImport.ps1 that defines target Dataverse table 
and field names.

.PARAMETER LookupData
Array of all lookup entries for this table, used to find parent entries in hierarchical structures.

.PARAMETER CodeMaskSeparators
Array of separator characters used in the hierarchy (e.g., '.', '\', '/'). 
If not provided, defaults to @('.', '\', '/').

.OUTPUTS
Returns a hashtable containing the Dataverse entity attributes, formatted for the data loader tool.
Includes the name field, optional description field, and parent lookup field for hierarchies.

.EXAMPLE
$entity = Convert-LookupValueToEntity -Item $lookupItem -LookupTableMapping $mapping -LookupData $allItems -CodeMaskSeparators @('.')

.NOTES
- For hierarchical lookups, the parent is determined by removing the last segment after the separator
- The parent lookup reference includes value, lookupentity, and lookupentityname properties
#>
function Convert-LookupValueToEntity {
    param(
        $Item,
        [hashtable]$LookupTableMapping,
        [array]$LookupData,
        [array]$CodeMaskSeparators = @()
    )
    $attributes = @{}

    $attributes[$LookupTableMapping.DataverseTable.NameField] = Invoke-TextFieldHandler -Value $Item.FullValue
    if ($LookupTableMapping.DataverseTable.DescriptionField) {
        $attributes[$LookupTableMapping.DataverseTable.DescriptionField] = Invoke-TextFieldHandler -Value $Item.Description
    }
    
    #hierarchy mapping
    if ($LookupTableMapping.MapAsHierarchy -and $LookupTableMapping.DataverseTable.PrevLevelLookupField -and $Item.FullValue) {
        # Determine which separators to use
        $separatorsToUse = if ($CodeMaskSeparators -and $CodeMaskSeparators.Count -gt 0) {
            $CodeMaskSeparators
        } else {
            @('.', '\', '/')  # Fallback to defaults
        }
        
        # Find the last occurrence of any separator in the FullValue
        $lastSeparatorIndex = -1
        foreach ($separator in $separatorsToUse) {
            $index = $Item.FullValue.LastIndexOf($separator)
            if ($index -gt $lastSeparatorIndex) {
                $lastSeparatorIndex = $index
            }
        }
        
        # If we found a separator, this entry has a parent
        if ($lastSeparatorIndex -gt 0) {
            $parentFullValue = $Item.FullValue.Substring(0, $lastSeparatorIndex)
            
            # Find the parent entry in LookupData by matching FullValue
            $parentEntry = $LookupData | Where-Object { $_.FullValue -eq $parentFullValue } | Select-Object -First 1
            
            if ($parentEntry) {
                $attributes[$LookupTableMapping.DataverseTable.PrevLevelLookupField] = @{
                    value            = $parentEntry.Id
                    lookupentity     = $LookupTableMapping.DataverseTable.TableLogicalName
                    lookupentityname = $parentEntry.FullValue
                }
            }
        }
    }

    return $attributes
}
