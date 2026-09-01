. $PSScriptRoot\Core.ps1
. $PSScriptRoot\TableOperations.ps1
. $PSScriptRoot\Defaults.ps1

<#
.SYNOPSIS
Helper function to invoke Dataverse metadata API requests.

.DESCRIPTION
Wrapper function for making GET requests to the Dataverse metadata API.
#>
function Invoke-DataverseHttpGet {
    param (
        [Parameter(Mandatory)]
        [string]$query
    )
    
    $uri = $baseURI + $query
    $headers = $baseHeaders.Clone()
    $headers.Add('If-None-Match', $null)
    $headers.Add('Prefer', 'odata.include-annotations="*"')
    
    $request = @{
        Uri     = $uri
        Method  = 'Get'
        Headers = $headers
    }
    
    Invoke-ResilientRestMethod $request
}

<#
.SYNOPSIS
Loads lookup table data from Dataverse into memory for field mapping during import.

.DESCRIPTION
The Get-LookupTableData function retrieves lookup table data as specified in the 
Defaults.ps1. The data is loaded from the connected Dataverse 
environment and stored in a hashtable structure for efficient lookup during project and resource import.

.EXAMPLE
$lookupData = Get-LookupTableData
# Access RBS lookup data
$rbsData = $lookupData["RBS"]
# Find a specific RBS value (bottom level only)
$rbsRecord = $rbsData["Corporate.Finance - Management.Finance Team"]

# Get all hierarchical levels for proper parent-child relationship handling
$allLevels = Get-HierarchicalLookupLevels -FullValue 'Corporate.Quality - Management.Quality Team' -LookupName 'RBS' -LookupDataCache $lookupData

.OUTPUTS
Returns a hashtable where:
- Keys are the lookup reference names (e.g., "RBS", "Health")
- Values are nested hashtables mapping FullValue to record data (ID, Name, Description)
#>

function Get-LookupTableData {
    param()
    
    Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
    Write-Host "Loading Lookup Table Data from Dataverse" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor Cyan
    
    $lookupDataCache = @{}
    
    foreach ($lookupTable in $LookupTablesToLoad) {
        $lookupName = $lookupTable.Lookup
        $dataverseTable = $lookupTable.DataverseTable
        
        Write-Host "`nLoading lookup table: $lookupName" -ForegroundColor Yellow
        
        $lookupRecords = @{}
        
        # Load flat lookup data from a single table
        $tableCollectionName = $dataverseTable.TableCollectionName
        
        # Query all records from this table
        $query = "?`$select=$($dataverseTable.NameField)"
        
        try {
            $response = Get-Records -setName $tableCollectionName -query $query
            $records = $response.value
            
            Write-Host "    Retrieved $($records.Count) records" -ForegroundColor Gray
            
            foreach ($record in $records) {
                $recordId = $record."$($dataverseTable.TableLogicalName)id"  # add 'id'
                $recordName = $record.$($dataverseTable.NameField)
                
                # For flat lookups, use the name as the key
                $lookupRecords[$recordName] = @{
                    Id                  = $recordId
                    Name                = $recordName
                    TableCollectionName = $tableCollectionName
                }
            }
        } catch {
            Write-Host "    Error loading data from $tableCollectionName : $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $lookupDataCache[$lookupName] = $lookupRecords
        Write-Host "  Cached $($lookupRecords.Count) total records for lookup '$lookupName'" -ForegroundColor Green
    }
    
    Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
    Write-Host "Lookup table data loading completed" -ForegroundColor Cyan
    Write-Host "Loaded $($lookupDataCache.Count) lookup table(s)" -ForegroundColor Cyan
    Write-Host "--------------------------------------------`n" -ForegroundColor Cyan
    
    return $lookupDataCache
}

<#
.SYNOPSIS
Loads choice field options from Dataverse into memory for field mapping during import.

.DESCRIPTION
The Get-ChoiceFieldData function retrieves choice field options (picklist values) from Dataverse.
It supports both local choice fields (specific to an entity/table) and global choice fields (shared across entities).

For local choices, it queries the entity metadata to get option values.
For global choices, it queries the global option set metadata.

.EXAMPLE
$choiceData = Get-ChoiceFieldData
# Access a choice field's options
$statusOptions = $choiceData["ProjectStatus"]
# Find a specific option value by label
$activeOption = $statusOptions | Where-Object { $_.Label -eq 'Active' }
# Use: $Project['cr62c_status'] = $activeOption.Value

.OUTPUTS
Returns a hashtable where:
- Keys are the choice field reference names (e.g., "ProjectStatus")
- Values are arrays of hashtables containing option metadata (Value, Label, Description)
#>

function Get-ChoiceFieldData {
    param()
    
    Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
    Write-Host "Loading Choice Field Data from Dataverse" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor Cyan
    
    $choiceDataCache = @{}
    
    foreach ($choiceField in $ChoiceFieldsToLoad) {
        $choiceName = $choiceField.Name
        $isGlobalChoice = $choiceField.IsGlobalChoice
        $dataverseTable = $choiceField.DataverseTable
        
        Write-Host "`nLoading choice field: $choiceName" -ForegroundColor Yellow
        Write-Host "  Type: $(if ($isGlobalChoice) { 'Global Choice' } else { 'Local Choice' })" -ForegroundColor Gray
        
        $options = @()
        
        try {
            if ($isGlobalChoice) {
                # For global choices, query the global option set metadata
                Write-Host "  Loading global choice from entity: $($dataverseTable.TableLogicalName), attribute: $($dataverseTable.ChoiceField)" -ForegroundColor Gray
                
                # Query: /api/data/v9.2/EntityDefinitions(LogicalName='entityname')/Attributes(LogicalName='attributename')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=GlobalOptionSet($select=Options)
                $metadataQuery = "EntityDefinitions(LogicalName='$($dataverseTable.TableLogicalName)')/Attributes(LogicalName='$($dataverseTable.ChoiceField)')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?`$select=LogicalName&`$expand=GlobalOptionSet(`$select=Options)"
                
                $response = Invoke-DataverseHttpGet -query $metadataQuery
                
                if ($response.GlobalOptionSet -and $response.GlobalOptionSet.Options) {
                    foreach ($option in $response.GlobalOptionSet.Options) {
                        $optionValue = $option.Value
                        $optionLabel = $option.Label.UserLocalizedLabel.Label
                        $optionDescription = if ($option.Description.UserLocalizedLabel) { $option.Description.UserLocalizedLabel.Label } else { $null }
                        
                        $options += @{
                            Value       = $optionValue
                            Label       = $optionLabel
                            Description = $optionDescription
                        }
                    }
                }
            } else {
                # For local choices, query the entity attribute metadata
                Write-Host "  Loading local choice from entity: $($dataverseTable.TableLogicalName), attribute: $($dataverseTable.ChoiceField)" -ForegroundColor Gray
                
                # Query: /api/data/v9.2/EntityDefinitions(LogicalName='entityname')/Attributes(LogicalName='attributename')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)
                $metadataQuery = "EntityDefinitions(LogicalName='$($dataverseTable.TableLogicalName)')/Attributes(LogicalName='$($dataverseTable.ChoiceField)')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?`$select=LogicalName&`$expand=OptionSet(`$select=Options)"
                
                $response = Invoke-DataverseHttpGet -query $metadataQuery
                
                if ($response.OptionSet -and $response.OptionSet.Options) {
                    foreach ($option in $response.OptionSet.Options) {
                        $optionValue = $option.Value
                        $optionLabel = $option.Label.UserLocalizedLabel.Label
                        $optionDescription = if ($option.Description.UserLocalizedLabel) { $option.Description.UserLocalizedLabel.Label } else { $null }
                        
                        $options += @{
                            Value       = $optionValue
                            Label       = $optionLabel
                            Description = $optionDescription
                        }
                    }
                }
            }
            
            Write-Host "    Retrieved $($options.Count) options" -ForegroundColor Gray
        } catch {
            Write-Host "    Error loading choice field data: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $choiceDataCache[$choiceName] = $options
        Write-Host "  Cached $($options.Count) options for choice '$choiceName'" -ForegroundColor Green
    }
    
    Write-Host "`n--------------------------------------------" -ForegroundColor Cyan
    Write-Host "Choice field data loading completed" -ForegroundColor Cyan
    Write-Host "Loaded $($choiceDataCache.Count) choice field(s)" -ForegroundColor Cyan
    Write-Host "--------------------------------------------`n" -ForegroundColor Cyan
    
    return $choiceDataCache
}

<#
.SYNOPSIS
Gets the lookup value for flat or self-referential hierarchical lookup tables.

.DESCRIPTION
The Get-LookupValue function retrieves the record information for a single lookup value
from either a flat lookup table or a self-referential hierarchical table. Unlike
Get-HierarchicalLookupLevels which returns all levels across multiple tables, this function
returns a single record from a single table.

Use this function for:
- Flat lookups (non-hierarchical single table lookups)
- Self-referential hierarchies (hierarchical lookups stored in a single table)

.PARAMETER Value
The lookup value to find. For flat lookups, this is the simple value (e.g., 'Active').
For self-referential hierarchies, this is the full path (e.g., 'Corporate.Quality - Management.Quality Team').

.PARAMETER LookupName
The lookup reference name (e.g., 'RBS', 'Health') as defined in LookupAndChoiceFieldsForMapping.ps1

.PARAMETER LookupDataCache
The lookup data cache returned from Get-LookupTableData function

.EXAMPLE
$lookupData = Get-LookupTableData
$healthValue = Get-LookupValue -Value 'Active' -LookupName 'Health' -LookupDataCache $lookupData

# Returns:
# @{ Id = '...'; Name = 'Active'; TableCollectionName = 'cr62c_healths'; Description = '...' }

.EXAMPLE
$lookupData = Get-LookupTableData
$rbsValue = Get-LookupValue -Value 'Corporate.Quality - Management.Quality Team' -LookupName 'RBS' -LookupDataCache $lookupData

# Returns:
# @{ Id = '...'; Name = 'Corporate.Quality - Management.Quality Team'; TableCollectionName = 'cr62c_rbshierarchies'; Description = '...' }

.OUTPUTS
Returns a hashtable containing:
- Id: The Dataverse record ID
- Name: The name/value of the lookup item
- TableCollectionName: The Dataverse table collection name
- Description: The description field value (if available)
Returns $null if the value is not found.
#>

function Get-LookupValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,
        
        [Parameter(Mandatory = $true)]
        [string]$LookupName,

        [Parameter(Mandatory = $false)]
        [bool]$ReturnIDOnly = $false
    )
    
    # Get the lookup records for this lookup name
    if (-not $global:lookupData.ContainsKey($LookupName)) {
        Write-Host "Warning: Lookup '$LookupName' not found in cache" -ForegroundColor Yellow
        return $null
    }
    
    $lookupRecords = $global:lookupData[$LookupName]
    
    # Check if the value exists in the cache
    if (-not $lookupRecords.ContainsKey($Value)) {
        Write-Host "Warning: Value '$Value' not found in lookup '$LookupName'" -ForegroundColor Yellow
        return $null
    }
    
    # Return the record
    $record = $lookupRecords[$Value]
    
    # return @{
    #     Id = $record.Id
    #     Name = $record.Name
    #     TableCollectionName = $record.TableCollectionName
    # }

    if (-not $record) {
        return $null
    }
    
    if ($ReturnIDOnly) {
        return $record.Id
    }
    #format ready for populating into an insert or update in Dataverse
    return "/$($record.TableCollectionName)($($record.Id))"
}

<#
.SYNOPSIS
    Get a choice field option value by label
    
.DESCRIPTION
    Retrieves the numeric option value for a choice field based on the label text.
    Searches the loaded choice data cache for a matching label (case-insensitive).
    
.PARAMETER Label
    The display text of the choice option to find
    
.PARAMETER ChoiceName
    The name of the choice field (as defined in the configuration)
    
.PARAMETER ChoiceDataCache
    The hashtable containing loaded choice field data
    
.PARAMETER ReturnFullOption
    If specified, returns the full option hashtable instead of just the Value
    
.EXAMPLE
    $statusValue = Get-ChoiceValue -Label "In Progress" -ChoiceName "ProjectStatus" -ChoiceDataCache $global:choiceData
    
.NOTES
    Returns null if the label is not found
#>
function Get-ChoiceValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Label,
        
        [Parameter(Mandatory = $true)]
        [string]$ChoiceName
    )
    
    # Handle empty or null values
    if ([string]::IsNullOrWhiteSpace($Label)) {
        return $null
    }
    
    # Check if the choice exists in the cache
    if (-not $global:choiceData.ContainsKey($ChoiceName)) {
        Write-Host "Warning: Choice field '$ChoiceName' not found in cache" -ForegroundColor Yellow
        return $null
    }
    
    # Get the options for this choice field
    $options = $global:choiceData[$ChoiceName]
    
    # Find the option with matching label (case-insensitive)
    $matchingOption = $options | Where-Object { $_.Label -eq $Label }
    
    if (-not $matchingOption) {
        Write-Host "Warning: Label '$Label' not found in choice field '$ChoiceName'" -ForegroundColor Yellow
        return $null
    }
    
    return $matchingOption.Value
}