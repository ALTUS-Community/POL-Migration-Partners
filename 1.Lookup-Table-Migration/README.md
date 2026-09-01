# Script overview

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working lookup-table scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

This script is intended to be as part of a migration of data from Project Online to Altus. This particular script is OPTIONAL and handles the export of Project Online Lookup Table values and subsequent import of this data into an Altus environment. This is intended as a one-off single run migration of data.

The working scripts and migration tools are supplied separately through the Altus Partner Portal. Do not unblock or execute the files in this repository.

## Prerequisites

The script requires:

- [PowerShell 7.4 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)

The user running the script must have sufficient access to the Project Online environment (source) and Dataverse environment (destination) in order to read and write the required Lookup Table data.

This script also requires PnP.PowerShell library to be installed. (Install-Module PnP.PowerShell)

## Running the Scripts

To run the script, open Powershell or SharePoint Online Management Shell and navigate to the location of the script. (Note: If the version of Powershell you are running is lower than 7 then PowerShell 7 will be called and executed from within that window).

Two distinct scripts are required - performing the Export and Import operations as separately scripted actions.

### Export Lookup Table Values

Select to run the following command;

```powershell
.\ExportLookupTableData.ps1
```

The script can also optionally be run by passing in the following parameters:

- PwaUrl - The URL to the Project Online PWA environment which contains the lookup tables that you wish to export data from e.g. "https://tenant.sharepoint.com/sites/PWA"
- ClientId - Optional Entra ID Client ID for authentication. Defaults to Sensei app client ID.
- OutputFolder - Sets the destination folder that Lookup Table. (Default is otherwise for data be stored in a folder called LookupTables relative to where script is run from)

Example call with parameters:

```powershell
.\ExportLookupTableData.ps1 -PwaUrl "https://tenant.sharepoint.com/sites/PWA/" -ClientId "YOUR-ENTRA-CLIENT-ID" -OutputFolder "C:\Temp\LookupTables\"
```

If any parameters are not passed in to the call to the script, then the defaults will be used as defined at the top of the script file.

### Import Lookup Table Values

Values from Project Online lookup tables can be migrated to Altus in two different modes.

1. Flat Table - When migrating as a flat table, any hierarchical structure present in the source Project Online lookup table will be flattened.
2. Self Referential Hierarchy Table - When migrating as a self referential hierarchy table, hierarchical data will be populated into the table with the parent value referred to in a lookup field.

Destination tables in Dataverse that are intended to store the Lookup Table values MUST be created before attempting to run the import script. This is a manual process as each organisation will have its own rules in terms of publisher prefixes, table naming conventions and solution/environment lifecycle. Tables cannot be created via the Import script - only population of data into those already existing tables.

A table that is being created for contain Lookup Table values should contain the following field(s):

- Name (this field should be created automatically when creating a new table in Dataverse)
- Description (optional)

For a self referential lookup table, it will in addition need the following field:

- Parent (name this field as you would like) which should be a Lookup field which references back to the table you are creating the field for.

Once the Parent field is created, navigate to the Relationships list for the table. Select the relationship that was created when the Lookup field was created. (The relationship should be FROM your table TO your table). Select the _Hierarchy_ checkbox and _Save_ the relationship.

There are also other aspects to consider when creating Tables in Dataverse for use within Altus. Users of Altus who need to interact with those tables must be granted permissions to those tables via a security role and depending on requirements, additional security roles may be required to be configured in Altus as additional security roles to be given to Team/Group owners of Projects.

For single select values, any tables created for migration of Project Online Lookup Table values should also be referenced as lookup fields for whichever Altus artefact that will need to reference those lookup values in Altus. (For example, a Lookup Table in Project Online that acts as a Project level custom field table should be set up in Altus as a table which is then referenced in a Lookup field on the Project table in Dataverse (and then optionally also exposed to the UI in the Project Form)).

Project Online lookup fields can be configured to allow multiple selections. Dataverse does not allow a table's lookup field to contain multiple selections (although Choice fields can be configured that way). If you intend to maintain a multi select relationship, then rather than creating a Dataverse lookup field, you should instead;

- Ensure the table has been created in Dataverse to store the available lookup values
- Then, from the primary entity (e.g. Project or Bookable Resource), create a _many-to-many_ relationship to the table that contains the lookup table values
- Optionally, a many-to-many relationship can be exposed to the primary table's Form UI using the Subgrid component

#### Configure Lookup Table to Dataverse Table Mappings

The next step is to edit and save the LookupTablesToImport.ps1 file to configure the mapping between the exported Project Online lookup table data that you wish to import into Altus and the Dataverse tables that you have created to store that data.

$LookupTableFileLocation can be updated to a different location if the default location was not used when running the ExportLookupTableData.ps1 script.

$LookupTableMappings contains an array of mappings. An array item should be created for each Lookup table you wish to migrate the data for.
Examples are provided (commented out) including an example for a Flat Table mapping and another for a Self Referential Hierarchy Table mapping. Schema for the mappings are defined as follows:

SourceFile: Should contain the name of the generated CSV file exported from ExportLookupTableData.ps1 (e.g. LookupTable_Health.csv)
MapAsHierarchy: A value of $false maps to a flat table. A value of $true maps to a self referential hierarchy table.  
DataverseTable: An object containing the following properties:

- TableLogicalName: the Dataverse table LOGICAL (singular) name to map to
- TableCollectionName: the Dataverse table COLLECTION (plural) name to map to
- NameField: Field in Dataverse table to map the Name/Value to
- DescriptionField: (Optional) Field in Dataverse table to map the Description to
- PrevLevelLookupField: (Required only when MapAsHierarchy = $true) Identifies the self referential lookup field on that table. NOTE: Field name is case sensitive

#### Run ImportLookupTableData script

Ensure you have edited the LookupTablesToImport.ps1 file (as above) to identify the lookup data you wish to import.

Then, select to run the following command;

```powershell
./ImportLookupTableData.ps1 -D365Url "https://orgname.crm.dynamics.com/" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe"
```

The script can also optionally be run by passing in the following parameters:

- D365Url - The URL of the Dataverse environment that you wish to import the lookup table data to
- MigrationExePath - Required path to the separately approved `Altus.DevOps.D365.DataMigration.exe`; it is not bundled in this repository
- Force - Boolean value indicating whether existing records should be updated ($true = update existing)
- ParallelRequests - The maximum number of parallel requests that the data loader should undertake
- EnableDisablingOfPlugins - Indicates whether disabling of plugins is required in order for the import to succeed

The repository does not run this script. A separately approved working copy must provide both the target URL and the migration executable path.

Example call with parameters:

```powershell
./ImportLookupTableData.ps1 -D365Url "https://orgname.crm.dynamics.com/" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe"
```

If the D365 parameter is not passed in to the call to the script, then the user will be prompted to enter it at runtime.

During the running of the script, the script will generate a Data.xml file in a _DataLoaderLookupTables_ folder relative to the folder the script is run from. That file will then be used in the Import process. Logs will be written to the _Logs_ folder relative to the folder the script is run from and will contain a log for ImportLookupTableData_yyyymmddhhmmss.txt and also DataMigration-yyyymmdd.txt (note that the latter will be appended with new data if the script is run multiple times on the same day).

After running the script (assuming no issues were encountered) the mapped Dataverse tables should now contain the mapped Lookup Table data.
