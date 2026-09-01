[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are used by dot-sourcing scripts')]
param()
# CONFIGURATION
$ProjectDesktopExternalSystemId = 'REPLACE-WITH-TARGET-EXTERNAL-SYSTEM-ID'
if ($ProjectDesktopExternalSystemId -eq 'REPLACE-WITH-TARGET-EXTERNAL-SYSTEM-ID') {
    throw 'Project Desktop external system mapping contains a placeholder. Configure target record IDs in a working copy before use.'
}

# LOOKUP TABLES TO LOAD FOR REFERENCE
# Note: These lookup tables must be pre-loaded before running the Task Field migration to ensure proper mapping of lookup field values.
$LookupTablesToLoad = @(
    # (Examples below)
    @{
        Lookup         = "TaskScheduleKPI" # Name of the lookup table that you will use internally to this script to identify the dataset
        DataverseTable = @{
            TableLogicalName    = "cr9f9_taskschedulekpi"
            TableCollectionName = "cr9f9_taskschedulekpis"
            NameField           = "cr9f9_name"
        }
    },
    @{
        Lookup         = "TaskWorkKPI" # Name of the lookup table that you will use internally to this script to identify the dataset
        DataverseTable = @{
            TableLogicalName    = "cr9f9_taskworkkpi"
            TableCollectionName = "cr9f9_taskworkkpis"
            NameField           = "cr9f9_name"
        }
    }
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
# Note: These choice fields must be pre-loaded before running the Task Field migration to ensure correct mapping of choice field values.
$ChoiceFieldsToLoad = @(
    # @{
    #     Name = "Cost Type" # Name of the choice field that you will use internally to this script to identify the dataset
    #     IsGlobalChoice = $false  # Set to $true if this is a global choice field, otherwise $false for local choice fields
    #     DataverseTable = @{
    #         TableLogicalName = "sensei_task"     # Logical/internal name of the Dataverse table
    #         ChoiceField = "cr62c_costtype"  # Logical/internal name of the choice field in the Dataverse table
    #     }
    # },
    # @{
    #     Name = "Investment Category" # Name of the choice field that you will use internally to this script to identify the dataset
    #     IsGlobalChoice = $true   # Set to $true if this is a global choice field, otherwise $false for local choice fields
    #     DataverseTable = @{
    #         TableLogicalName = "sensei_task"     # Logical/internal name of the Dataverse table
    #         ChoiceField = "sensei_investmentcategory"  # Logical/internal name of the choice field in the Dataverse table
    #     }
    # }
)

