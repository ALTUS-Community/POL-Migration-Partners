[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are used by dot-sourcing scripts')]
param()
# CONFIGURATION
$EnableProjectNameFallback = $false  # Set to $true if projects with matching names have been pre-created in the Dataverse environment (use with caution to avoid incorrect matches)

# RESOURCE DEFAULTS
$DefaultPrimaryRoleId = 'REPLACE-WITH-TARGET-ROLE-ID'
$DefaultEnterpriseCalendarId = 'REPLACE-WITH-TARGET-CALENDAR-ID'
$DefaultTargetUtilisation = 100

# PROJECT DEFAULTS
$ProjectDesktopExternalSystemId = 'REPLACE-WITH-TARGET-EXTERNAL-SYSTEM-ID'
$DefaultProjectTypeId = 'REPLACE-WITH-TARGET-PROJECT-TYPE-ID'

if ($DefaultPrimaryRoleId -like 'REPLACE-WITH-*' -or
    $DefaultEnterpriseCalendarId -like 'REPLACE-WITH-*' -or
    $ProjectDesktopExternalSystemId -like 'REPLACE-WITH-*' -or
    $DefaultProjectTypeId -like 'REPLACE-WITH-*') {
    throw 'Migration defaults contain placeholders. Configure target record IDs in a working copy before use.'
}

# LOOKUP TABLES TO LOAD FOR REFERENCE
# Note: These lookup tables must be pre-loaded before running the Project and Resource migration to ensure proper mapping of lookup field values.
$LookupTablesToLoad = @(
    # (Examples below)
    # @{
    #     Lookup = "RBS" # Name of the lookup table that you will use internally to this script to identify the dataset
    #     DataverseTable = @{
    #         TableLogicalName = "cr62c_rbshierarchy"
    #         TableCollectionName = "cr62c_rbshierarchies"
    #         NameField = "cr62c_name"
    #     }
    # },
    # @{
    #     Lookup = "ChangeImpact" # Name of the lookup table that you will use internally to this script to identify the dataset
    #     DataverseTable = @{
    #         TableLogicalName = "cr62c_organisationalchangeimpact"
    #         TableCollectionName = "cr62c_organisationalchangeimpacts"
    #         NameField = "cr62c_name"
    #     }
    # },
    # @{
    #     Lookup = "Department" # Name of the lookup table that you will use internally to this script to identify the dataset
    #     DataverseTable = @{
    #         TableLogicalName = "cr62c_department"
    #         TableCollectionName = "cr62c_departments"
    #         NameField = "cr62c_name"
    #     }
    # },
    # @{
    #     Lookup = "Skills" # Name of the lookup table that you will use internally to this script to identify the dataset
    #     DataverseTable = @{
    #         TableLogicalName = "cr62c_skills"
    #         TableCollectionName = "cr62c_skillses"
    #         NameField = "cr62c_name"
    #     }
    # }
)

# CHOICE FIELDS TO LOAD FOR REFERENCE
# Note: These choice fields must be pre-loaded before running the Project and Resource migration to ensure correct mapping of choice field values.
$ChoiceFieldsToLoad = @(
    # @{
    #     Name = "Cost Type" # Name of the choice field that you will use internally to this script to identify the dataset
    #     IsGlobalChoice = $false  # Set to $true if this is a global choice field, otherwise $false for local choice fields
    #     DataverseTable = @{
    #         TableLogicalName = "sensei_bookableresource"     # Logical/internal name of the Dataverse table
    #         ChoiceField = "cr62c_costtype"  # Logical/internal name of the choice field in the Dataverse table
    #     }
    # },
    # @{
    #     Name = "Investment Category" # Name of the choice field that you will use internally to this script to identify the dataset
    #     IsGlobalChoice = $true   # Set to $true if this is a global choice field, otherwise $false for local choice fields
    #     DataverseTable = @{
    #         TableLogicalName = "sensei_project"     # Logical/internal name of the Dataverse table
    #         ChoiceField = "sensei_investmentcategory"  # Logical/internal name of the choice field in the Dataverse table
    #     }
    # }
)

