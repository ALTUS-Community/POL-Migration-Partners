# Script overview

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working task-field scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

This script is intended to be used as part of a migration of data from Project Online to Altus. This particular script handles the population of Task Custom Field data into an Altus environment based on exported Project Online data.

It is important that this script is run _after_ the Projects have been migrated to Altus _and after_ the schedule migration steps have been performed. (If run out of order, this script will not succeed because the Project's tasks will not yet be present in Altus).

The working scripts and migration tools are supplied separately through the Altus Partner Portal. Do not unblock or execute the files in this repository.

## Prerequisites

The script requires:

- [PowerShell 7.4 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [Azure CLI](https://aka.ms/installazurecliwindows) — used to authenticate to your Dataverse environment. When the script runs, it will open a browser window prompting you to sign in with the account that has access to the target Altus environment. If switching between environments, you will be prompted to sign in again.

The script relies on Project Online data having been extracted from a Project Online environment using either the [Microsoft export script](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online) or the ExportAllProjects.ps1 script in the 0.Project-Online-Extraction folder.

The data extracted from Project Online which the user intends to migrate into Altus must then be placed in the /Files folder relative to the location of _this_ script. _In addition_ ensure that you also copy the {ProjectName}\_Published_mpp.xml files that should have been generated when running the Schedule-Migration script.

The user running the script must have sufficient access to the Dataverse environment in order to read and write the required Project and Task data. Ideally they should be an Altus Admin User.

## Running the Script

Before running the script, ensure you copy the exported Project Online data files that have been used by these migration scripts that you wish to import into your Altus environment to the 'Files' folder adjacent to this script. (If that folder does not exist, go ahead and create it.)

Ensure that when copying these from the /Files folder in the 3.Schedule-Migration directory, that you copy the files from /Published back to sit at the same level as the other export files when you copy them to the ./5.TaskField-Migration/Files directory. Also ensure that you copy the series of {ProjectName}\_Published_mpp.xml files that should have been generated during the Schedule-Migration script activities.

To run the script, open Powershell or SharePoint Online Management Shell and navigate to the location of the script. (Note: If the version of Powershell you are running is lower than 7 then PowerShell 7 will be called and executed from within that window).

Select to run the following command;

```powershell
.\AltusTaskFieldMigration.ps1
```

The script can also optionally be run by passing in the following parameters:

- EnvironmentUrl - The URL to the environment which contains your Altus instance e.g. `https://orgname.crm.dynamics.com/`
- ExecutionMode - Allows the user to select between WhatIf mode or Execution mode (see below for further details on Execution Mode). Valid values are "W" or "WhatIf" for What-If mode or "E" or "Execute" for Execute mode.
- ForceLogin - This optional flag can be added to force a login prompt to be presented to the user. Can be used to switch between users when accessing the same environment.

Example call with parameters:

```powershell
.\AltusTaskFieldMigration.ps1 -EnvironmentUrl "https://orgname.crm.dynamics.com/" -ExecutionMode W
```

If any parameters are not passed in to the call to the script, then the user will be prompted to enter them interactively.

### Execution Mode

The script can be run in What-If or Execution mode. When running the script, the user will be prompted for their execution mode preference.

What-If mode will run the script but will not execute any commands that actually create data in your Dataverse environment. The proposed changes however will be logged.

Execute mode will action the changes and update records in your Dataverse environment.

Each time the script runs, a log file will be created.

The script can be run in Execute mode multiple times in an environment.

## Task Field Import

Task Field import will perform the following actions:

- Read all Published MPP files from the /Files folder relative to the script
- For each Published MPP file;
  - Look for an existing Altus Project which is linked to the Project Desktop external system for the project
  - Locate the related ...\_reporting_Tasks.json file
  - Locate the related ...\_published.xml file
  - Locate the related ...\_published_mpp.xml file
  - Read the ...\_reporting_Tasks.json to identify the Tasks for the Project
  - For each Task found in the ...\_reporting_Tasks.json file;
    - Determine the sensei_taskid of the related Altus task using the ...\_published.xml and ...\_published_mpp.xml files
    - Populate field values for the Task based on the script configuration in ImportTasks.ps1
    - Update the Tasks (if in Execution mode)
  - If in What-If mode, no Task updates will be performed. Instead, a log entry will be noted.

### Troubleshooting

#### Dataverse Authentication timeouts

If the Task Field Import operation is being run for a large number of tasks it can happen that the Dataverse authentication token expires while still attempting to process tasks for import. If this occurs, you will likely start to see logs which include the error 'Response status code does not indicate success: 401 (Unauthorized)'.

If this occurs, it is advisable to remove the associated files of the successfully imported tasks out of the /Files directory and to then restart the script to begin processing the task imports from where it left off.

## Custom Field Mapping

The scripts provide a mechanism to map Task fields from Project Online into Dataverse columns.

### Custom Field Mapping Prerequisites

1. The destination column must already exist in Dataverse on the appropriate table:
   - Tasks: `sensei_task`
2. Use the **logical name** (lowercase) of the Dataverse column, not the display name.

### Lookup Table and Choice Field Value Mapping

If the Task field you intend to save to is of type Lookup or of type Choice - or if the data is related to Lookup table data via a many-to-many relationship (to allow multiple selections) - then the script must pre-load the available values for that column into memory in order to correctly identify the value(s) to save. Tables that contain lookup data for a many-to-many relationship should be configured as though they were Lookup Tables in the configuration below.

You can configure which Lookup Tables and Choice values to pre-load via the _Defaults.ps1_ file

Lookup Tables should be identified in the $LookupTablesToLoad variable under the _# LOOKUP TABLES TO LOAD FOR REFERENCE_ heading and should include the following properties in each element in the hashtable:

- Lookup - Name of the lookup table that you will use internally to this script to identify the dataset
- DataverseTable - An object which includes the following properties;
  - TableLogicalName - The logical name of the table that contains the lookup table data
  - TableCollectionName - The collection name (plural) of the table that contains the lookup table data
  - NameField - The logical name of the field on the lookup table that contains the Name value

Choice OptionSets should be identified in the $ChoiceFieldsToLoad variable under the _# CHOICE FIELDS TO LOAD FOR REFERENCE_ heading and should include the following properties in each element in the hashtable;

- Name - Name of the choice field that you will use internally to this script to identify the dataset
- IsGlobalChoice - Set to $true if this is a global choice field, otherwise $false for local choice fields
- DataverseTable - An object which includes the following properties;
  - TableLogicalName - Logical/internal name of the Dataverse table that contains the reference to this choice
  - ChoiceField - Logical/internal name of the choice field in the Dataverse table

### Where to Edit

- **Tasks:** `ImportTasks.ps1` — edit the **USER CONFIGURATION (EDIT HERE)** section at the top of the file (above the `===== Don't edit below this line =====` marker).

### Data Sources

| Entity | File Pattern             | Where Fields Live                                                                            |
| ------ | ------------------------ | -------------------------------------------------------------------------------------------- |
| Tasks  | `*_reporting_Tasks.json` | OOTB: direct properties on `ReportingProjectTasksData.Tasks`; Custom: `CustomFields[]` array |
| Tasks  | `*_published.xml`        | Used by the script to derive the Altus Task ID                                               |
| Tasks  | `*_published_mpp.xml`    | Used by the script to derive the Altus Task ID                                               |

### Mapping OOTB (Out-of-the-Box) Fields

OOTB fields are direct properties on the task object. Uncomment and edit lines inside the mapping functions.

**Tasks** — edit `Add-TaskDataverseFieldMappings` in `ImportTasks.ps1`:

```powershell
    # $Task['cr_task_text_ootb'] = $TaskName
    # $Task['cr_task_datetime_ootb'] = ($ReportingTask.TaskEarlyFinish -as [datetime])
    # $Task['cr_task_dateonly_ootb'] = ($ReportingTask.TaskEarlyFinish -as [datetime]).ToString("yyyy-MM-dd")
    # $Task['cr_task_whole_ootb'] = ($ReportingTask.TaskPercentWorkCompleted -as [int])
    # $Task['cr_task_decimal_ootb'] = ($ReportingTask.TaskTotalSlack -as [decimal])
```

### Mapping Enterprise Custom Fields (Tasks)

Project Online Enterprise Custom Fields are exported under `ReportingTaskData.Task.CustomFields[]`. Each entry has:

- `CustomFieldName` — the display name of the field in Project Online.
- `CustomFieldValue.'#text'` — the actual value (as a string).

Use the helper function `Get-ReportingTaskCustomFieldTextValue` inside `Add-TaskDataverseFieldMappings`:

```powershell
# Text field
$textValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName 'Your Text Field'
if ($textValue) { $Task['cr_task_text_custom'] = [string]$textValue }

# Date field
$dateValue = Get-ReportingTaskCustomFieldTextValue -ReportingTask $ReportingTask -CustomFieldName 'Your Date Field'
if ($dateValue) { $Task['cr_task_datetime_custom'] = ($dateValue -as [datetime]) }

# Lookup fields can be matched using the Set-TaskLookupField function
# Note: To map to a Dataverse lookup field, you must ensure the table data is pre-loaded by this script in the Defaults.ps1 file
#       so that the correct corresponding lookup value can be populated.
# IMPORTANT: -DataverseFieldName is CASE SENSITIVE when identifying the Dataverse Schema Name for the lookup field
# Also - if your incoming Project Online custom field is a multi select lookup but you wish to only map to a Dataverse lookup field (that allows only single selection), then using this function will use the _first_ selected value for that multi select lookup for the task field.
Set-TaskLookupField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Task Schedule KPI' -LookupName 'TaskScheduleKPI' -DataverseFieldName 'cr9f9_TaskScheduleKPI'

# Choice fields can be matched using the Set-TaskChoiceField function
# Note: To map to a Dataverse choice field, you must ensure the choice field data is pre-loaded by this script in the Defaults.ps1 file
#       so that the correct corresponding choice value can be populated.
# Use the -IsMulti flag to denote when you are intending to save Project Online multi select lookup values to a Dataverse multi select choice field.
Set-TaskChoiceField -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Investment Category' -ChoiceName 'Investment Category' -DataverseFieldName 'sensei_investmentcategory'

# Many-To-Many Relationship Lookups can be matched using the Set-TaskManyToManyRelationship function
# Note: To map to a Dataverse many-to-many relationship, you must ensure the related table data is pre-loaded by this script in the Defaults.ps1 file
#       so that the correct corresponding related records can be populated.
# In order to populate the many-to-many relationship table, you must identify the some key aspects of the relationship table including;
# -RelationshipName - the name of the relationship as defined in Dataverse
# -RelatedTableCollectionName - the table Collection Name (plural) of the related table
# -RelatedTablePrimaryKey - the primary key (id) field of the related table
Set-TaskManyToManyRelationship -TaskBody $Task -ReportingTask $ReportingTask -POLCustomFieldName 'Task Work KPI' -LookupName 'TaskWorkKPI' -RelationshipName 'cr9f9_TaskWorkKPI_sensei_task_sensei_task' -RelatedTableCollectionName 'cr9f9_taskworkkpis' -RelatedTablePrimaryKey 'cr9f9_taskworkkpiid'
```

### Supported Data Types

| Dataverse Type   | PowerShell Cast           | Example                                                             |
| ---------------- | ------------------------- | ------------------------------------------------------------------- |
| Text             | `[string]`                | `$Task['cr_text'] = [string]$value`                                 |
| Whole Number     | `-as [int]`               | `$Task['cr_int'] = ($value -as [int])`                              |
| Decimal/Currency | `-as [decimal]`           | `$Task['cr_dec'] = ($value -as [decimal])`                          |
| Date and Time    | `-as [datetime]`          | `$Task['cr_datetime'] = ($value -as [datetime])`                    |
| Date Only        | `.ToString("yyyy-MM-dd")` | `$Task['cr_date'] = ($value -as [datetime]).ToString("yyyy-MM-dd")` |
| Lookup           | See below                 | See below                                                           |
| Choice           | See below                 | See below                                                           |

> **Important:** Dataverse has two date column types with different formats:
>
> - **Date and Time** (`Edm.DateTimeOffset`) — accepts full datetime values via `-as [datetime]`
> - **Date Only** (`Edm.Date`) — requires a string in `yyyy-MM-dd` format
>
> Using `-as [datetime]` on a Date Only field will cause an error like: _"Cannot convert the literal '...' to the expected type 'Edm.Date'"_

#### Lookup Fields

To successfully populate a lookup field, ensure first that you are pre-loading the available data options by defining
the lookup tables you wish to map values for in the Defaults.ps1 configuration.

Use the helper function Set-TaskLookupField to correctly identify the matching value in the Altus Dataverse table. These functions automatically handle the `@odata.bind` syntax and Id references required by Dataverse.

#### Choice Fields

To successfully populate a choice field, ensure first that you are pre-loading the available data options by defining the choice option sets that you wish to map values for in the Defaults.ps1 configuration.

Use the helper functions Set-TaskChoiceField to correctly identify the matching value in the Dataverse option set. These functions automatically handle the substitution of the numerical value required by Dataverse.

> **Important:** Certain task level fields in Altus should NOT be updated via this script and are explicitly excluded from any
> update operations that are generated from this script.  
> These include:
>
> - sensei_duration
> - sensei_effort
> - sensei_effortcompleted
> - sensei_effortremaining
> - sensei_effortremainingoverride
> - sensei_externalparenttaskid
> - sensei_externalprojectid
> - sensei_externaltaskid
> - sensei_fieldlock
> - sensei_formupdated
> - sensei_hash
> - sensei_index
> - sensei_isexternal
> - sensei_json_properties
> - sensei_lastsyncdate
> - sensei_level
> - sensei_locked
> - sensei_milestone
> - sensei_name
> - sensei_notes
> - sensei_order
> - sensei_originofchange
> - sensei_parenttask
> - sensei_percentcomplete
> - sensei_percentcompletedecimal
> - sensei_project
> - sensei_scheduletype
> - sensei_summary
> - sensei_taskfinish
> - sensei_taskfinishdatetime
> - sensei_tasknumber
> - sensei_taskstart
> - sensei_taskstartdatetime
> - sensei_tasktype
