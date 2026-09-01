[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are used by dot-sourcing scripts')]
param()

# ------------------------------
# USER CONFIGURATION (EDIT HERE)
# ------------------------------
# Edit below the lookup tables that you wish to import and identify the Dataverse table that you wish to map to.
# Lookup Table data can be migrated either as a flat table (any hierarchy levels will be flattened) or as a hierarchy self referential lookup table
# The destination table MUST already exist in Dataverse with the appropriate fields created. (See README)

$LookupTableFileLocation = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "LookupTables")

$LookupTableMappings = @(
  #   EXAMPLE FLAT MAPPING
  # @{
  #   SourceFile       = "LookupTable_Department.csv"                     # CSV file exported from ExportLookupTableData.ps1
  #   MapAsHierarchy   = $false                                           # A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.
  #   DataverseTable   = @{ TableLogicalName     = "cr62c_department"     # Dataverse table LOGICAL (singular) name to map to.
  #                         TableCollectionName  = "cr62c_departments"    # Dataverse table COLLECTION (plural) name to map to. 
  #                         NameField            = "cr62c_name"           # Field in Dataverse table to map the Name/Value to
  #                       }
  # },
  # @{
  #   SourceFile       = "LookupTable_Health.csv"                         # CSV file exported from ExportLookupTableData.ps1
  #   MapAsHierarchy   = $false                                           # A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.
  #   DataverseTable   = @{ TableLogicalName     = "cr62c_health"         # Dataverse table LOGICAL (singular) name to map to.
  #                         TableCollectionName  = "cr62c_healths"        # Dataverse table COLLECTION (plural) name to map to. 
  #                         NameField            = "cr62c_name"           # Field in Dataverse table to map the Name/Value to
  #                         DescriptionField     = "cr62c_description"    # (Optional) Field in Dataverse table to map the Description to
  #                       }
  # },
  # @{
  #   SourceFile       = "LookupTable_Organisational Change Impact.csv"   # CSV file exported from ExportLookupTableData.ps1
  #   MapAsHierarchy   = $false                                           # A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.
  #   DataverseTable   = @{ TableLogicalName     = "cr62c_organisationalchangeimpact"        # Dataverse table LOGICAL (singular) name to map to.
  #                         TableCollectionName  = "cr62c_organisationalchangeimpacts"       # Dataverse table COLLECTION (plural) name to map to. 
  #                         NameField            = "cr62c_name"           # Field in Dataverse table to map the Name/Value to
  #                       }
  # },
  #   EXAMPLE SELF REFERENTIAL HIERARCHY MAPPING
  # @{
  #   SourceFile       = "LookupTable_Skills.csv"                         # CSV file exported from ExportLookupTableData.ps1
  #   MapAsHierarchy   = $true                                            # A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.
  #   DataverseTable   = @{ TableLogicalName     = "cr62c_skills"         # Dataverse table LOGICAL (singular) name to map to.
  #                         TableCollectionName  = "cr62c_skillses"       # Dataverse table COLLECTION (plural) name to map to.
  #                         NameField            = "cr62c_name"           # Field in Dataverse table to map the Name/Value to
  #                         PrevLevelLookupField = "cr62c_parentskill"    # PrevLevelLookup property should identify the self referential lookup field for this table. 
  #                       }       
  # },
  # @{
  #   SourceFile       = "LookupTable_RBS.csv"                            # CSV file exported from ExportLookupTableData.ps1
  #   MapAsHierarchy   = $true                                            # A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.
  #   DataverseTable   = @{ TableLogicalName     = "cr62c_rbshierarchy"   # Dataverse table LOGICAL (singular) name to map to.
  #                         TableCollectionName  = "cr62c_rbshierarchies" # Dataverse table COLLECTION (plural) name to map to.
  #                         NameField            = "cr62c_name"           # Field in Dataverse table to map the Name/Value to
  #                         DescriptionField     = "cr62c_description"    # (Optional) Field in Dataverse table to map the Description to
  #                         PrevLevelLookupField = "cr62c_parentrbs"      # PrevLevelLookup property should identify the self referential lookup field for this table. 
  #                       }       
  # }
)
