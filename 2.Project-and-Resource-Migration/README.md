# Script overview

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working project and resource scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

This script is intended to be used as part of a migration of data from Project Online to Altus. This particular script handles the creation of Resources and Projects in an Altus environment based on exported Project Online data.

The working scripts and migration tools are supplied separately through the Altus Partner Portal. Do not unblock or execute the files in this repository.

## Prerequisites

The script requires:

- [PowerShell 7.4 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [Azure CLI](https://aka.ms/installazurecliwindows) — used to authenticate to your Dataverse environment. When the script runs, it will open a browser window prompting you to sign in with the account that has access to the target Altus environment. If switching between environments, you will be prompted to sign in again.

The script relies on Project Online data having been extracted from a Project Online environment using either the [Microsoft export script](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online) or the ExportAllProjects.ps1 script in the 0.Project-Online-Extraction folder.

The data extracted from Project Online which the user intends to migrate into Altus must then be placed in the /Files folder relative to the location of _this_ script.

The user running the script must have sufficient access to the Dataverse environment in order to read and write the required Resource and Project data. Ideally they should be an Altus Admin User.

For Importing Projects, part of the process requires updating the MPP files to include custom parameters. In order to Import Projects, you will need Microsoft Project installed locally.

## Running the Script

Before running the script, ensure you copy the exported data files from the Project Online Extraction script that you wish to import into your Altus environment to the 'Files' folder adjacent to this script. (If that folder does not exist, go ahead and create it.)

To run the script, open Powershell or SharePoint Online Management Shell and navigate to the location of the script. (Note: If the version of Powershell you are running is lower than 7 then PowerShell 7 will be called and executed from within that window).

Select to run the following command;

```powershell
.\AltusPOLMigration.ps1
```

The script can also optionally be run by passing in the following parameters:

- EnvironmentUrl - The URL to the environment which contains your Altus instance e.g. https://orgname.crm.dynamics.com/
- ExecutionMode - Allows the user to select between WhatIf mode or Execution mode (see below for further details on Execution Mode). Valid values are "W" or "WhatIf" for What-If mode or "E" or "Execute" for Execute mode.
- Action - This represents the action that should be taken by the script. Options are "ImportNamedResources", "ImportNamedAndGenericResources" or "ImportProjects".
- ForceLogin - This optional flag can be added to force a login prompt to be presented to the user. Can be used to switch between users when accessing the same environment.

Example call with parameters:

```powershell
.\AltusPOLMigration.ps1 -EnvironmentUrl "https://orgname.crm.dynamics.com/" -ExecutionMode W -Action ImportNamedAndGenericResources
```

If any parameters are not passed in to the call to the script, then the user will be prompted to enter them interactively.

### Execution Mode

The script can be run in What-If or Execution mode. When running the script, the user will be prompted for their execution mode preference.

What-If mode will run the script but will not execute any commands that actually create data in your Dataverse environment. The proposed changes however will be logged.

Execute mode will action the changes and create records in your Dataverse environment.

Each time the script runs a log file will be created.

The script can be run in Execute mode multiple times in an environment. If data from a previous run already exists, duplicate records will not be created.

## Resource Import

There are two options for Resource Import in this script. 'Named Only' or 'Generic and Named'.

- Read all existing Named Bookable Resources in your Dataverse environment
- Read all existing System Users in your Dataverse environment
- Read all of the Project Online resources json files that were extracted from your Project Online environment, deduplicate them and consolidate them according to type
- For each unique Project Online Named resource, the script will;
  - Look for a matching Named Bookable Resource (based on the login name of the Project Online user).
    - If a match is found, the Resource properties will be updated and a log entry will be noted
  - Look for a matching System User (based on the login name of the Project Online user)
    - If no match is found no action will be taken and a log entry will be noted
  - If there is no matching Named Bookable Resource and a matching System User is found
    - If in Execute mode, a Named Bookable Resource will be created
      - Values for Primary Role, Target Utilization and Enterprise Calendar will be set as per the values set in the Defaults.ps1 file
    - If in What-If mode, no Resource will be created in Dataverse, but a log entry will be noted
- When run in 'Generic and Named' mode, the script will also;
  - Look for a matching Generic Bookable Resource (based on the name of the Project Online generic resource)
    - If a match is found, no action will be taken and a log entry will be noted
  - If there is no matching Generic Bookable Resource;
    - If in Execute mode, a Generic Bookable Resource will be created
    - If in What-If mode, no Resource will be created in Dataverse, but a log entry will be noted

## Project Import

Project import will perform the following actions:

- Read all Project Desktop aligned Projects from the Altus environment
- Read the Solution version number for the Altus Atsumeru Solution in the Altus environment
- Read the Organization name from the environment
- Get the EnvironmentId for the environment (via Power Platform Admin API)
- Read all Published MPP files from the /Files folder relative to the script
- For each Published MPP file;
  - Read the ProjectName and ProjectGUID from the related .json file
  - Look for an existing Altus Project which is linked to the Project Desktop external system for the project
    - If a match is found, the Project properties will be updated and a log entry will be noted
  - Look for an existing Altus Project which has a matching External Project ID property but no matching External System record
    - If a match is found, the Project properties will be updated and an ExternalProject record will be created in Altus (and aligned to the Project Desktop external system)
  - (OPTIONAL) Fallback to Match on the Name of an existing Project which does not have an External Project ID property or an External System record
    - The fallback option is disabled by default, but can be configured by editing the Defaults.ps1 file to set $EnableProjectNameFallback to $true
    - When matching on Project Name, the match must be exact (but is case insensitive - so 'Project One' will match on 'project one')
    - Differences in whitespace characters will not be treated as a match - so 'Project One' will not match to ' Project One', 'Project One' or 'Project One '
    - If a match is found, the Project properties will be updated and an ExternalProject record will be created in Altus (and aligned to the Project Desktop external system)
  - If there is no matching Project in Altus;
    - If in Execute mode;
      - A Project will be created in Altus
        - The Project Type will be set as per the value set in the Defaults.ps1 file
      - An associated ExternalProject record will be created in Altus (and aligned to the Project Desktop external system)
  - The Published MPP file in the Files directory will be updated to include the custom properties required to align it to the Altus Project when used with the Altus for Project add-in. If the custom properties already contain the correct values, no changes will be made.
  - If in What-If mode, no Project, External Project or Custom Property updates will be performed. Instead, a log entry will be noted.

### Troubleshooting

#### Microsoft Project timeouts

During the Project Import operation where the Published MPP files are updated to include custom properties, it is advisable to monitor the Microsoft Project window as some files will trigger a dialog box to be shown to the user asking for their action. If the dialog is not dismissed/actioned by the user then the custom property update will time out and the script will log an error for that .mpp update. (The script will then continue with the remaining projects, but the .mpp for the project which timed out will not be updated).

_If any time outs are encountered during the process it is advisable to re-run the Import Project script for those affected Projects ensuring that the Microsoft Project window is monitored to interact with any blocking dialog panes._

#### Dataverse Authentication timeouts

If the Project Import operation is being run for a large number of projects it can happen that the Dataverse authentication token expires while still attempting to process projects for import. If this occurs, you will likely start to see logs which include the error 'Response status code does not indicate success: 401 (Unauthorized)'.

If this occurs, it is advisable to remove the associated files of the successfully imported projects out of the /Files directory and to then restart the script to begin processing the project imports from where it left off.

## Custom Field Mapping

The scripts provide a mechanism to map additional Project and Resource fields from Project Online into Dataverse columns.

### Custom Field Mapping Prerequisites

1. The destination column must already exist in Dataverse on the appropriate table:
   - Projects: `sensei_project`
   - Resources: `sensei_bookableresource`
2. Use the **logical name** (lowercase) of the Dataverse column, not the display name.

### Lookup Table and Choice Field Value Mapping

If the field you intend to save to (for either a Project or a Resource) is of type Lookup or of type Choice - or if the data is related to Lookup table data via a many-to-many relationship (to allow multiple selections) - then the script must pre-load the available values for that column into memory in order to correctly identify the value(s) to save. Tables that contain lookup data for a many-to-many relationship should be configured as though they were Lookup Tables in the configuration below.

You can configure which Lookup Tables and Choice values to pre-load via the _Defaults.ps1_ file

Lookup Tables should be identified in the $LookupTablesToLoad variable under the _# LOOKUP TABLES TO LOAD FOR REFERENCE_ heading and should include the following properties in each element in the hashtable:

- Lookup - Name of the lookup table that you will use internally to this script to identify the dataset
- DataverseTable - An object which includes the following properties;
  - TableLogicalName - The logical name of the table that contains the lookup table data
  - TableCollectionName - The collection name (plural) of the table that contains the lookup table data
  - NameField - The logical name of the field on the lookup table that contains the Name value

Choice OptionSets should be identified in the $ChoiceFieldsToLoad variable under the _# CHOICE FIELDS TO LOAD FOR REFERENCE_ heading and should include the following properties in each element in the hashtable;

- Name - Name of the choice filed that you will use internally to this script to identify the dataset
- IsGlobalChoice - Set to $true if this is a global choice field, otherwise $false for local choice fields
- DataverseTable - An object which includes the following properties;
  - TableLogicalName - Logical/internal name of the Dataverse table that contains the reference to this choice
  - ChoiceField - Logical/internal name of the choice field in the Dataverse table

### Where to Edit

- **Projects:** `ImportProjects.ps1` — edit the **USER CONFIGURATION (EDIT HERE)** section at the top of the file (above the `===== Don't edit below this line =====` marker).
- **Resources:** `ImportResources.ps1` — edit the **USER CONFIGURATION (EDIT HERE)** section at the top of the file (above the `===== Don't edit below this line =====` marker).

### Data Sources

| Entity    | File Pattern                 | Where Fields Live                                                                         |
| --------- | ---------------------------- | ----------------------------------------------------------------------------------------- |
| Projects  | `*_reporting.json`           | OOTB: direct properties on `ReportingProjectData.Project`; Custom: `CustomFields[]` array |
| Resources | `*_reporting_Resources.json` | Direct properties on each resource object                                                 |

### Mapping OOTB (Out-of-the-Box) Fields

OOTB fields are direct properties on the project/resource object. Uncomment and edit lines inside the mapping functions.

**Projects** — edit `Add-ProjectDataverseFieldMappings` in `ImportProjects.ps1`:

```powershell
$Project['cr_project_text_ootb'] = $ProjectName
$Project['cr_project_datetime_ootb'] = ($ReportingProject.ProjectStartDate -as [datetime])
$Project['cr_project_dateonly_ootb'] = ($ReportingProject.ProjectStartDate -as [datetime]).ToString("yyyy-MM-dd")
$Project['cr_project_whole_ootb'] = ($ReportingProject.ProjectIdentifier -as [int])
$Project['cr_project_decimal_ootb'] = ($ReportingProject.ProjectCalendarDuration -as [decimal])
```

**Resources** — edit `Add-NamedResourceDataverseFieldMappings` or `Add-GenericResourceDataverseFieldMappings` in `ImportResources.ps1`:

```powershell
$ResourceBody['cr_resource_text_ootb'] = $ProjectResource.ResourceName
$ResourceBody['cr_resource_datetime_ootb'] = ($ProjectResource.ResourceCreatedDate -as [datetime])
$ResourceBody['cr_resource_dateonly_ootb'] = ($ProjectResource.ResourceCreatedDate -as [datetime]).ToString("yyyy-MM-dd")
$ResourceBody['cr_resource_whole_ootb'] = ($ProjectResource.ResourceType -as [int])
$ResourceBody['cr_resource_decimal_ootb'] = ($ProjectResource.ResourceStandardRate -as [decimal])
```

### Mapping Enterprise Custom Fields (Projects)

Project Online Enterprise Custom Fields are exported under `ReportingProjectData.Project.CustomFields[]`. Each entry has:

- `CustomFieldName` — the display name of the field in Project Online.
- `CFValue.'#text'` — the actual value (as a string).

Use the helper function `Get-ReportingCustomFieldTextValue` inside `Add-ProjectDataverseFieldMappings`:

```powershell
# Text field
$textValue = Get-ReportingCustomFieldTextValue -ReportingProject $ReportingProject -CustomFieldName 'Your Text Field'
if ($textValue) { $Project['cr_project_text_custom'] = [string]$textValue }

# Date field
$dateValue = Get-ReportingCustomFieldTextValue -ReportingProject $ReportingProject -CustomFieldName 'Your Date Field'
if ($dateValue) { $Project['cr_project_date_custom'] = ($dateValue -as [datetime]) }

# Lookup fields can be matched using the Set-ProjectLookupField function
# Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding lookup value can be populated.
# IMPORTANT: -DataverseFieldName is CASE SENSITIVE when identifying the schema name of your Dataverse lookup field
# Also - if your incoming Project Online custom field is a multi select lookup but you wish to only map to a Dataverse lookup field (that allows only single selection), then using this function will use the _first_ selected value for that multi select lookup for the project field.
Set-ProjectLookupField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Organisational Change Impact' -LookupName 'ChangeImpact' -DataverseFieldName 'cr62c_OrganisationalChangeImpact'

# Choice fields can be matched using the Set-ProjectChoiceField function
# Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding choice value can be populated.
# Use the -IsMulti flag to denote when you are intending to save Project Online multi select lookup values to a Dataverse multi select choice field.
Set-ProjectChoiceField -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Investment Category' -ChoiceName 'Investment Category' -DataverseFieldName 'sensei_investmentcategory'

# Many-To-Many Relationship Lookups can be matched using the Set-ProjectManyToManyRelationship function
# Note: To map to a Dataverse many-to-many relationship table, you must ensure the related table's data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding values can be populated.
# In order to populate the many-to-many relationship table, you must identify the some key aspects of the relationship table including;
# -RelationshipName - the name of the relationship as defined in Dataverse
# -RelatedTableCollectionName - the table Collection Name (plural) of the related table
# -RelatedTablePrimaryKey - the primary key (id) field of the related table
Set-ProjectManyToManyRelationship -ProjectBody $Project -ReportingProject $ReportingProject -POLCustomFieldName 'Multi Department' -LookupName 'Department' -RelationshipName 'cr62c_sensei_project_cr62c_Department_cr62c_Department' -RelatedTableCollectionName 'cr62c_departments' -RelatedTablePrimaryKey 'cr62c_departmentid'
```

### Mapping Enterprise Custom Fields (Resources)

Project Online Enterprise Custom Fields are exported under `ProjectResource.CustomFields[]`. Each entry has:

- `CustomFieldName` — the display name of the field in Project Online.
- `CustomFieldValue.'#text'` — the actual value (as a string).

Use the helper function `Get-ResourceCustomFieldTextValue` inside `Add-NamedResourceDataverseFieldMappings` and/or `Add-GenericResourceDataverseFieldMappings`:

```powershell
# Text field - Note Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
$textValue = Get-ResourceCustomFieldTextValue -ProjectResource $projectResource -CustomFieldName 'Your Text Field'
if ($textValue) { $ResourceBody['cr_resource_text_custom'] = [string]$textValue }

# Date field - Note Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
$dateValue = Get-ResourceCustomFieldTextValue -ProjectResource $projectResource -CustomFieldName 'Your Date Field'
if ($dateValue) { $ResourceBody['cr_project_date_custom'] = ($dateValue -as [datetime]) }

# Lookup fields can be matched using the Set-ResourceLookupField function
# Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding lookup value can be populated.
# Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
# IMPORTANT: -DataverseFieldName is CASE SENSITIVE when identifying the schema name of your Dataverse lookup field
# Also - if your incoming Project Online custom field is a multi select lookup but you wish to only map to a Dataverse lookup field (that allows only single selection), then using this function will use the _first_ selected value for that multi select lookup for the resource field.
Set-ResourceLookupField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'RBS' -LookupName 'RBS' -DataverseFieldName 'cr62c_RBSHierarchy'

# Choice fields can be matched using the Set-ResourceChoiceField function
# Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding choice value can be populated.
# Add the -IsMulti parameter and set it to $true when you are intending to save Project Online multi select lookup values to a Dataverse _multi select choice_ field.
# Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
Set-ResourceChoiceField -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'Cost Type' -ChoiceName 'Cost Type' -DataverseFieldName 'cr62c_costtype'

# Many-To-Many Relationship Lookups can be matched using the Set-ResourceManyToManyRelationship function
# Note: To map to a Dataverse many-to-many relationship table, you must ensure the related table's data is pre-loaded by this script in the Defaults.ps1 file so that the correct corresponding values can be populated.
# In order to populate the many-to-many relationship table, you must identify the some key aspects of the relationship table including;
# -RelationshipName - the name of the relationship as defined in Dataverse
# -RelatedTableCollectionName - the table Collection Name (plural) of the related table
# -RelatedTablePrimaryKey - the primary key (id) field of the related table
# Substitute $projectResource for $genericResource if in the Add-GenericResourceDataverseFieldMappings function
Set-ResourceManyToManyRelationship -ResourceBody $ResourceBody -ProjectResource $ProjectResource -POLCustomFieldName 'Multi Skills' -LookupName 'Skills' -RelationshipName 'cr62c_sensei_bookableresource_cr62c_Skills_cr62c_Skills' -RelatedTableCollectionName 'cr62c_skillses' -RelatedTablePrimaryKey 'cr62c_skillsid'
```

### Supported Data Types

| Dataverse Type   | PowerShell Cast           | Example                                                                |
| ---------------- | ------------------------- | ---------------------------------------------------------------------- |
| Text             | `[string]`                | `$project['cr_text'] = [string]$value`                                 |
| Whole Number     | `-as [int]`               | `$project['cr_int'] = ($value -as [int])`                              |
| Decimal/Currency | `-as [decimal]`           | `$project['cr_dec'] = ($value -as [decimal])`                          |
| Date and Time    | `-as [datetime]`          | `$project['cr_datetime'] = ($value -as [datetime])`                    |
| Date Only        | `.ToString("yyyy-MM-dd")` | `$project['cr_date'] = ($value -as [datetime]).ToString("yyyy-MM-dd")` |
| Lookup           | See below                 | See below                                                              |
| Choice           | See below                 | See below                                                              |

> **Important:** Dataverse has two date column types with different formats:
>
> - **Date and Time** (`Edm.DateTimeOffset`) — accepts full datetime values via `-as [datetime]`
> - **Date Only** (`Edm.Date`) — requires a string in `yyyy-MM-dd` format
>
> Using `-as [datetime]` on a Date Only field will cause an error like: _"Cannot convert the literal '...' to the expected type 'Edm.Date'"_

#### Lookup Fields

To successfully populate a lookup field, ensure first that you are pre-loading the available data options by defining
the lookup tables you wish to map values for in the Defaults.ps1 configuration.

Use the helper functions Set-ProjectLookupField or Set-ResourceLookupField to correctly identify the matching value in the Altus Dataverse table. These functions automatically handle the `@odata.bind` syntax and Id references required by Dataverse.

#### Choice Fields

To successfully populate a choice field, ensure first that you are pre-loading the available data options by defining the choice option sets that you wish to map values for in the Defaults.ps1 configuration.

Use the helper functions Set-ProjectChoiceField or Set-ResourceChoiceField to correctly identify the matching value in the Dataverse option set. These functions automatically handle the substitution of the numerical value required by Dataverse.

## Task Custom Field Mapping

Task-level custom field values from Project Online are imported via the **Altus for Project add-in** during the MPP publish process, not by this PowerShell script. This allows you to map individual task custom fields to Altus task columns.

### How It Works

1. PowerShell script creates the project and project tasks in Altus
2. User opens the MPP file in Microsoft Project with the Altus for Project add-in installed
3. User publishes the MPP to Altus via the add-in
4. During publish, the add-in reads configured custom field mappings from **projectDesktopConfig** and populates task columns in Altus

### Task Custom Field Mapping Prerequisites

1. Custom columns must already exist in Altus on the `sensei_task` table
2. Custom fields must be configured in the MPP file (project-local or enterprise-level)
3. You must have access to configure **projectDesktopConfig** in your Altus environment (requires Altus Admin or configuration privileges)

### Identifying Custom Field Names

Refer to the [Microsoft Project Desktop custom fields reference](https://support.microsoft.com/en-au/office/custom-fields-in-project-desktop-604eaea9-9154-491a-9c00-764e5d46603e) to identify which field slot your custom field uses.

| Field Type | Slot Name                                           | Example   |
| ---------- | --------------------------------------------------- | --------- |
| Text       | `Text1` through `Text27` (`Text27-30` are reserved) | `Text7`   |
| Number     | `Number1` through `Number20`                        | `Number3` |
| Date       | `Date1` through `Date9` (`Date10` is reserved)      | `Date5`   |
| Flag       | `Flag1` through `Flag20`                            | `Flag2`   |

### Configuring Task Custom Field Mappings

Navigate to **Altus** → **Settings** → **Microsoft Project Configuration** in your Altus environment.

Under **Custom Field Mappings**, click **New Field Mapping**:

| Field                       | Value                        | Notes                                              |
| --------------------------- | ---------------------------- | -------------------------------------------------- |
| **Entity**                  | `sensei_task`                | Maps to task-level fields                          |
| **Microsoft Project Field** | `Text7` (or your field slot) | Use the underlying slot name, not the display name |
| **Altus Field**             | Select from dropdown         | Must be a field on the sensei_task table           |

**Example mappings:**

| Project Online Field                   | Slot    | Altus Column           |
| -------------------------------------- | ------- | ---------------------- |
| Strategic Risk Assessment (enterprise) | Text7   | sensei_risk_assessment |
| Project Priority (enterprise)          | Number2 | sensei_priority_score  |
| Deadline Milestone (enterprise)        | Date3   | sensei_milestone_date  |

### Supported Field Types

The add-in supports mapping the following Microsoft Project field types to Altus columns:

| MS Project Type                              | Altus Types Supported                          |
| -------------------------------------------- | ---------------------------------------------- |
| Text (Text1–Text26) (Text27-30 are reserved) | String, Memo, OptionSet, MultiSelect OptionSet |
| Number (Number1–Number20)                    | Integer, Decimal, Double, Money, OptionSet     |
| Date (Date1–Date9) (Date10 is reserved)      | Date/Time                                      |
| Duration                                     | Integer, Decimal, Double                       |
| Cost                                         | Money, Decimal, Integer                        |
| Flag (Flag1–Flag20)                          | Boolean, OptionSet (Yes/No only)               |

For detailed type-mapping rules and validation, see [Microsoft Project Configuration in Altus Docs](https://docs.altus.pro/products/AltusForProject/Configuration.html#custom-field-mappings).

### Important Notes

- **OptionSet mappings** use field **labels**, not values. Ensure field labels in MS Project exactly match OptionSet labels in Altus
- **MultiSelect OptionSets** require comma-separated values in MS Project (comma is the delimiter)
- Once a custom field is mapped, users must maintain the field definition exactly as configured — any changes to the field type or labels can break the mapping
- **Enterprise task custom fields are not currently supported** — only project-local custom fields can be mapped at this time

### Validating Your Mappings

When users publish the MPP file, the add-in will validate that:

1. The field exists in MS Project
2. The mapped Altus field exists
3. The field types are compatible

If validation fails, the add-in will display errors and the publish will not proceed. Review the error details and adjust the mapping in projectDesktopConfig, then try publishing again.
