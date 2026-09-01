# SharePoint to Dynamics 365 Migration Tool

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working SharePoint scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

This tool exports and imports SharePoint list data and documents to Microsoft Dynamics 365 using the MS CMT (Configuration Migration Tool) format. It works in conjunction with the Project Online and Project migration tools.

## Prerequisites

**Required:** Projects must already be present in your Altus (Dynamics 365) environment from a previous migration step. SharePoint list items will be linked to these projects during import.

**Note:** The POLExport data from previous migration steps (step 1) is used to identify and map projects in SharePoint.

**PowerShell Module Requirements:**

- **PnP.PowerShell** - The scripts will automatically attempt to install this module if not present
- If you do not have administrative rights or permissions to install PowerShell modules on your machine, please arrange for PnP.PowerShell to be installed by your IT administrator before running these scripts
- Manual installation command (if needed): `Install-Module -Name PnP.PowerShell -Scope CurrentUser`

## Migration Workflow & Prerequisites

> **ℹ️ NOTE:** The export scripts will only process projects that have associated project sites. Not every project in Project Online will have a project site, so some projects may be excluded from the SharePoint export process.

**This tool performs three steps:**

1. **Export Lists** – Export list data from SharePoint
   - Extract list items, metadata, and project references from configured SharePoint site collections
   - Convert to CMT XML format
   - Can use POLExport path from previous step 1 (optional)

2. **Import Lists** – Import list data to Dynamics 365
   - Requires projects to already exist in Altus (from step 1 migration)
   - Links list items to their corresponding projects in D365
   - Creates/updates records based on configuration

3. **Export Documents** – Download documents from SharePoint libraries
   - **NOT a full SharePoint migration** - this is document extraction only
   - Downloads documents from document libraries and preserves folder structure
   - Intended for use with SharePoint sync tools (OneDrive sync, etc.)

- Output files can be dragged into a SharePoint synced folder for upload
- Optional: Extended metadata and version history
- Simple mode (default): Fast download of files and folders

---

> ⚠️ **IMPORTANT: This Tool is Document Download Only**
>
> The **Export Documents** feature extracts files for local storage or sync workflows. It is **NOT a complete SharePoint site migration** and does **NOT include**:
>
> - Site structure, navigation, or pages
> - Permissions or sharing settings
> - Content types, metadata, or field definitions
> - Workflows or automations
> - Retention policies or compliance settings
> - Site columns or list templates
>
> **For Complete SharePoint Migration**, use enterprise tools:
>
> - **ShareGate** ⭐ Recommended - Full SharePoint site migration with all features
> - Microsoft SharePoint Migration Tool (SPMT)
> - Credeon, Sharegate, or other third-party migration platforms
>
> **This tool is designed to:** Accelerate document download so you can drag files into a SharePoint synced folder for manual upload to your target site.

---

## Configuration

All settings are defined in **`-Run-Migration.ps1`** within the `CONFIGURATION` section (lines 15–60).

### Site Collections

Define SharePoint site collections to migrate:

```powershell
$SiteCollections = @(
  @{
    Url           = "https://contoso.sharepoint.com/sites/pwa/"
    FolderName    = "PwaExample"         # Output folder name
    ProjectFilter = @()              # @() = all projects, or @("*2024*", "Project A")
  },
  @{
    Url           = "https://contoso.sharepoint.com/sites/example"
    FolderName    = "ExampleSite"
    ProjectFilter = @()
  }
)
```

**ProjectFilter** supports wildcard patterns:

- `@()` – Export all projects
- `@("*2024*")` – Export only projects matching "_2024_"
- `@("Project A", "Project B")` – Export specific projects

### Dynamics 365 Settings

```powershell
$D365Url = "https://orgname.crm.dynamics.com"  # Your D365 environment
$ImportParallelRequests = 4                             # Parallel threads (1–10)
$ImportForce = $false                                   # $true = update existing records
```

### POL Export Root (step 1 output)

```powershell
$PolExportRoot = ""  # Leave blank to default to ./POLExports next to -Run-Migration.ps1
```

- Point this to your existing step-1 POL export folder if it lives elsewhere (absolute path recommended)
- If left blank, the script uses `./POLExports/<FolderName>` beside `-Run-Migration.ps1`
- Each site collection entry still uses its `FolderName` to resolve the subfolder under this root

### Authentication (reference only)

```powershell
$ClientId = "YOUR-ENTRA-CLIENT-ID"
```

Authentication settings belong to the approved working tool. Do not add client secrets, consent URLs, or tenant-specific identifiers to this repository.

## Setup Requirements

Before running any migration scripts, you must configure two essential files:

### 1. Configure the Schema File (`data_schema.xml`)

The schema file defines the target Dynamics 365 entities and fields. **Maintain this file using the MS Configuration Migration Tool (CMT):**

1. Open the **Configuration Migration Tool** on your local machine
2. Load the existing `data_schema.xml` file
3. Modify entities and fields as needed to match your D365 customizations
4. Save the updated schema back to `data_schema.xml`
5. Verify that all entities and fields used in your field mappings are present

**Alternatively**, you can export a fresh schema directly from your D365 environment and replace the existing file.

**Why?** Using the CMT ensures your schema is valid and matches your D365 customizations, preventing validation errors during import.

### 2. Configure the Field Mapping File (`export.config.json`)

The mapping file controls how SharePoint fields are transformed to D365 attributes:

1. Open `export.config.json`
2. For each SharePoint list being migrated, add an entry to the `lists` array
3. Map each SharePoint field to its corresponding D365 attribute
4. Define value transformations (choiceMap, thresholds, multipliers, etc.)
5. Verify field names match the schema exactly (case-sensitive)

See the **Configuration Files** section below for detailed examples and syntax.

#### Copilot-assisted export.config.json drafting

You can have Copilot draft or refine `export.config.json` for you. Provide it with:

- This document (paste the relevant sections so Copilot knows the expected structure and field types)
- SharePoint list definitions from the REST API (`/_api/web/lists/getbytitle('{ListName}')/fields`) so it can see `InternalName`, data types, and choice values
- Your Altus entity definitions and attributes (from `data_schema.xml` or a Dataverse metadata export) so it can map to the correct logical names and option set values

Suggested prompt outline for Copilot:

1. "Here is the SharePoint list schema JSON from `_api/.../fields` for list <name>."
2. "Here is the current `export.config.json` snippet for this list (or leave blank if new)."
3. "Here are the Altus entity definitions/attributes for `<entity logical name>` including option set numeric values (see the Dataverse metadata GET below)."
4. "Draft/adjust the `columnMap` so each SharePoint field maps to the right D365 attribute. Use choiceMap for choices; use thresholds for numeric buckets; use lookup for people fields mapped to `systemuser`."

Review Copilot's output to ensure:

- `entityAttribute` names match your D365 environment exactly (case-sensitive)
- Option set integers match your Altus metadata
- `spFieldInternalName` values match the SharePoint REST response
- Any lookups target the correct entity (e.g., `systemuser` for people fields)

**Dataverse entity/attribute metadata (include for Copilot):**

Use these queries to retrieve entity definition and option set values:

**1. Get entity with all attributes:**

```
GET https://{your-org}.crm{region}.dynamics.com/api/data/v9.2/EntityDefinitions(LogicalName='{logicalName}')?$expand=Attributes($select=LogicalName,AttributeType,DisplayName),ManyToManyRelationships($select=SchemaName,IntersectEntityName),ManyToOneRelationships($select=ReferencingAttribute,ReferencingEntityNavigationPropertyName)&$select=DisplayName,EntitySetName,SchemaName,LogicalName,PrimaryIdAttribute,PrimaryNameAttribute
```

**2. Get option set values (for each picklist/choice field):**

```
GET https://{your-org}.crm{region}.dynamics.com/api/data/v9.2/EntityDefinitions(LogicalName='{logicalName}')/Attributes(LogicalName='{attributeLogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options),GlobalOptionSet($select=Options)
```

Run query #1 once per entity to get all attributes and their types. Then run query #2 for each OptionSet/Status attribute to retrieve the numeric values needed for choiceMap. Include both responses when asking Copilot to build mappings so it can:

- Confirm logical names match your schema
- Pull option set numeric values for choiceMap configurations
- Identify lookup fields and their target entities

## Quick Start - Two Options

### Option 1: Use the Migration Orchestrator (Recommended)

```powershell
.\-Run-Migration.ps1
```

This is the primary entry point. The script:

- Prompts you to choose: **Export Lists**, **Import Lists**, or **Export Documents**
- Displays configured site collections and D365 environment before starting
- Processes multiple SharePoint site collections sequentially
- Creates organized output folders for each site
- Provides detailed summary with success/failure tracking

Before running the orchestrator, set these in `-Run-Migration.ps1` (CONFIGURATION section):

- Site collections: `Url`, `FolderName`, and optional `ProjectFilter`
- POL export root: `$PolExportRoot` (leave blank to default to `./POLExports`; otherwise set your existing step-1 export path)
- Mapping path: `$MappingJsonPath` (export.config\*.json)
- Optional schema path: `$CmtSchemaPath` (data_schema\*.xml) - for reference only
- D365 import settings: `$D365Url`, `$ImportParallelRequests`, `$ImportForce`
- Client ID: `$ClientId` (configure only in the approved working tool)

Run the orchestrator once configured, and it will resolve per-site POL export paths as `$PolExportRoot/<FolderName>` and write outputs under `Output/<FolderName>/` (or `Output/Documents/<FolderName>/` for document export).

**Operation Modes:**

1. **Export Lists** - Extract SharePoint list data to CMT XML files
2. **Import Lists** - Load XML files into Dynamics 365
3. **Export Documents** - Download documents from SharePoint libraries

⚠️ **Important:** The document export feature downloads files only. It is **not a complete SharePoint migration**. For full SharePoint site migration including lists, pages, workflows, and site structure, use dedicated migration tools such as **ShareGate**.

### Option 2: Run Scripts Individually

Export, import, and document export are available as standalone scripts for advanced scenarios:

```powershell
# Step 1: Export SharePoint lists
.\1-export-lists.ps1 -SiteCollectionUrl "..." -OutputFolder ".\Output" -POLExportPath "..."

# Step 2: Import to Dynamics 365 (requires projects in Altus)
.\2-import-lists.ps1 -D365Url "..." -DataFolder ".\Output" -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe" -POLExportPath "..."

# Step 3: Export documents from SharePoint libraries
.\3-export-documents.ps1 -SiteCollectionUrl "..." -OutputFolder ".\Output\Documents" -POLExportPath "..."
```

Both the orchestrator and individual scripts require the same configuration settings defined below.

## Prerequisites & Permissions

### SharePoint Export

When running an export operation, **your user account must be a Site Collection Administrator** on each SharePoint site collection being exported. This is required because the export script needs to:

- Access all project sites and document libraries
- Read list items, metadata, and version history
- Retrieve property bag values

If you see an access denied error during export, ensure your account has Site Collection Admin privileges.

### Dynamics 365 Import

Your user account must have permissions to import data into the target Dynamics 365 environment.

### Authentication

Both SharePoint and Dynamics 365 use interactive authentication. You will be prompted to sign in during script execution.

## Core Scripts

### `1-export-lists.ps1` – SharePoint Export

Extracts list data from SharePoint and converts to CMT XML format.

**Called by:** `-Run-Migration.ps1` (automatically)

**Standalone usage:**

```powershell
.\1-export-lists.ps1 `
  -SiteCollectionUrl "https://contoso.sharepoint.com/sites/pwa/" `
  -MappingJsonPath ".\export.config.json" `
  -OutputFolder ".\Output\vNext" `
  -PolExportPath "C:\path\to\polexport\from\step\1"
```

_Optional:_ Add `-CmtSchemaPath ".\data_schema.xml"` if you want to use the schema file for validation reference.

**Parameters:**

- **-PolExportPath** (optional)
  - Path to the POLExport data from the previous Project Online extraction step (step 1)
  - Used to identify and map project references in SharePoint lists
  - If not provided, defaults to standard SharePoint list export without POL context

**Output:**

- `Output/<SiteName>/<ProjectName>/Data.xml` – CMT-compliant XML for each project

### `2-import-lists.ps1` – Dynamics 365 Import

Loads CMT XML files into D365 and links list items to projects already present in Altus.

**Prerequisite:** Projects must already exist in your Altus environment (from step 1 project migration)

**Called by:** `-Run-Migration.ps1` (automatically)

**Standalone usage:**

```powershell
.\2-import-lists.ps1 `
  -D365Url "https://orgname.crm.dynamics.com" `
  -DataFolder ".\Output\vNext" `
  -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe" `
  -POLExportPath "C:\path\to\polexport\from\step\1"
```

### `3-export-documents.ps1` – Document Export

Downloads documents from SharePoint document libraries for use with SharePoint sync tools or manual upload.

**⚠️ Important Note:**

- This script extracts **files only** from document libraries
- It is **NOT a complete SharePoint migration tool**
- No migration to another SharePoint environment is performed
- Exported files can be used with:
  - SharePoint sync (OneDrive sync) to move to a new location
  - Manual drag-and-drop into a synced folder
  - Backing up important documents
- **For complete SharePoint site migration** (lists, pages, workflows, permissions, etc.), use dedicated migration tools:
  - **ShareGate** (recommended for enterprise)
  - Microsoft's built-in SharePoint Migration Tool
  - Other third-party migration solutions

**Called by:** `-Run-Migration.ps1` (automatically when selecting "Export Documents")

**Standalone usage:**

```powershell
# Simple mode (default) - Just files and folders
.\3-export-documents.ps1 `
  -SiteCollectionUrl "https://contoso.sharepoint.com/sites/pwa/" `
  -OutputFolder ".\Output\Documents\vNext" `
  -POLExportPath "C:\path\to\polexport\from\step\1"

# Detailed mode - With metadata and version history
.\3-export-documents.ps1 `
  -SiteCollectionUrl "https://contoso.sharepoint.com/sites/pwa/" `
  -OutputFolder ".\Output\Documents\vNext" `
  -POLExportPath "C:\path\to\polexport\from\step\1" `
  -DetailedMetadata `
  -ExcludeVersionHistory:$false
```

**Parameters:**

- **-SiteCollectionUrl** (required)
  - SharePoint site collection URL

- **-OutputFolder** (optional)
  - Where to save downloaded documents
  - Default: `.\Output\Documents\Default`

- **-POLExportPath** (optional)
  - Path to POL export data from step 1
  - Used to identify and map project references

- **-ProjectFilter** (optional)
  - Array of wildcard patterns to filter projects
  - Example: `@("*2024*", "Project A")`

- **-DetailedMetadata** (switch, optional)
  - When enabled: Exports extended metadata (ETag, ContentType, UniqueId, all FieldValues)
  - When enabled: Creates folder per file with metadata.json
  - Default: $false (simple mode)

- **-ExcludeVersionHistory** (switch, optional)
  - When $true: Only exports current version
  - When $false: Exports all versions (requires -DetailedMetadata)
  - Default: $true

- **-ExcludeLibraries** (optional)
  - Comma-separated list of library names to skip
  - Default: `"Style Library,Preservation Hold Library,Form Templates,Recycle Bin,Site Assets"`

- **-ExcludeFilePatterns** (optional)
  - File patterns to exclude (wildcards supported)
  - Default: `@("*.aspx")`

- **-IncludeRootWeb** (switch, optional)
  - Include root web in project discovery
  - Default: $false

**Output Modes:**

**Simple Mode (default):**

```
Output/Documents/vNext/
├── Project A/
│   ├── Documents/
│   │   ├── file1.docx
│   │   ├── file2.pdf
│   │   ├── Subfolder/
│   │   │   └── file3.xlsx
│   │   ├── manifest.csv
│   │   └── manifest.json
│   └── Shared Documents/
│       └── ...
└── Project B/
    └── ...
```

Use with SharePoint Sync or manual upload:

1. Open OneDrive sync on your computer (Settings → OneDrive → Start sync)
2. Sync a folder in the new SharePoint location
3. Drag exported files from `Output/Documents/<Project>/Documents/` into the synced folder
4. OneDrive automatically uploads them to SharePoint

**Detailed Mode (`-DetailedMetadata`):**

```
Output/Documents/vNext/
├── Project A/
│   ├── Documents/
│   │   ├── file1.docx/
│   │   │   ├── file1.docx
│   │   │   ├── file1_v1.0.docx (if version history enabled)
│   │   │   └── metadata.json
│   │   ├── manifest.csv
│   │   └── manifest.json
│   └── ...
```

**Manifest Files:**

- Each library gets `manifest.csv` and `manifest.json` summarizing all exported documents
- Global manifests created at root: `documents-manifest.csv` and `documents-manifest.json`

---

## Configuration Files

### `export.config.json` – Field Mapping

Maps SharePoint fields to Dynamics 365 attributes. Controls data transformation and validation.

**Structure:**

```json
{
  "version": "1.0",
  "lists": [
    {
      "spListTitle": "Risks",
      "entityLogicalName": "sensei_risk",
      "projectLookupAttribute": "sensei_project",
      "backlinkAttribute": "sensei_sourceitemurl",
      "columnMap": [
        {
          "spFieldInternalName": "Status",
          "entityAttribute": "statuscode",
          "type": "Status",
          "choiceMap": {
            "(1) Active": 1,
            "(2) Postponed": 2
          }
        },
        {
          "spFieldInternalName": "Probability",
          "entityAttribute": "sensei_probability",
          "type": "OptionSet",
          "thresholds": {
            "0": 955000000,
            "0.2": 955000001,
            "0.4": 955000002,
            "0.6": 955000003,
            "0.8": 955000004
          }
        }
      ]
    }
  ]
}
```

**Key concepts:**

- **backlinkAttribute**: Optional field to store the SharePoint item URL in the Dynamics 365 record. This creates a link back to the original SharePoint item for reference after migration.

  **Use this only if:**
  - The SharePoint site will remain available after migration
  - The source items are in a permanent location that will persist
  - Users need to reference the original SharePoint items post-migration

  **Do not use if:**
  - Migrating away from Project Online (sites will be decommissioned)
  - SharePoint sites are temporary or will be deleted
  - The links would become broken over time

  If specified, the field must exist in your D365 schema. Leave empty (`""`) if not needed.

  ```json
  "backlinkAttribute": "sensei_sourceitemurl"
  ```

- **spFieldInternalName**: The internal field name in SharePoint (not the display name). To find this:
  1. Navigate to the SharePoint list
  2. Click the column dropdown and select "Edit column"
  3. Look at the URL in your browser's address bar
  4. Find the parameter `Field=` which contains the internal name (e.g., `Field=Status` or `Field=Bool_x0020_Test`)
  5. Use this value in your configuration

  **Alternative method:** Query the SharePoint REST API to retrieve the list definition and internal field names:

  ```
  https://{tenant}.sharepoint.com/sites/{SiteCollectionPath}/{WebPath}/_api/web/lists/getbytitle('{ListName}')/fields
  ```

  **URL tokens:**
  - `{tenant}` - Your SharePoint tenant name (e.g., "contoso" for contoso.sharepoint.com)
  - `{SiteCollectionPath}` - The site collection path after /sites/ (e.g., "vNext")
  - `{WebPath}` - The subsite path if accessing a subsite (omit if querying root site collection)
  - `{ListName}` - The display title of the SharePoint list (e.g., "Risks", "Issues")

  **Copilot-assisted setup:** If you provide Copilot with (1) the REST API list details above, (2) your current `export.config.json`, (3) optionally your `data_schema.xml` for reference, and (4) these instructions, it can draft the field mappings for you. OptionSet numeric values can be pulled by calling the Dataverse REST API for the entity definition attributes.
  For retrieving option set metadata via Dataverse REST API, see the Microsoft docs: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/retrieve-metadata-name-metadataid

  This returns all fields with their internal names, display names, and types. Look for the `InternalName` property in the response.

- **choiceMap**: Maps SharePoint choice values to D365 option set integers.

  ```json
  "choiceMap": { "(1) Active": 1, "(2) Postponed": 2 }
  ```

- **thresholds**: For numeric OptionSets (Probability, Impact), maps ranges to option values.

  ```json
  "thresholds": { "0": 955000000, "0.2": 955000001 }
  ```

  If a field value is 0.5, it matches threshold 0.2 (highest threshold ≤ value) → 955000001.

- **Field Types Supported:**
  - `Text` — Plain text field, no transformation needed
  - `Number`, `Money` — Numeric fields with optional `multiplier` or `divider` for scaling
    - Example – Scale a percentage to whole number:
      ```json
      {
        "spFieldInternalName": "Probability",
        "entityAttribute": "custom_percenttestwn",
        "type": "Number",
        "multiplier": 100
      }
      ```
    - Example – Divide a numeric value:
      ```json
      {
        "spFieldInternalName": "Cost",
        "entityAttribute": "custom_cost",
        "type": "Number",
        "divider": 1000
      }
      ```
  - `Decimal` — Decimal precision field with optional `multiplier`/`divider`
    - Example – Decimal field with precision:
      ```json
      {
        "spFieldInternalName": "ExchangeRate",
        "entityAttribute": "custom_exchangerate",
        "type": "Decimal",
        "multiplier": 1000
      }
      ```
  - `DateTime` — Full date and time with Z suffix (e.g., `2024-01-15T14:30:00Z`)
    - Example – DateTime field:
      ```json
      {
        "spFieldInternalName": "DueDate",
        "entityAttribute": "sensei_duedate",
        "type": "DateTime"
      }
      ```
  - `DateOnly` — Date with time component set to midnight (e.g., `2024-01-15T00:00:00`)
    - Example – DateOnly field:
      ```json
      {
        "spFieldInternalName": "StartDate",
        "entityAttribute": "sensei_startdate",
        "type": "DateOnly"
      }
      ```
  - `OptionSet`, `Status`, `State` — Use `choiceMap` or `thresholds` for value mapping
    ```json
    "choiceMap": { "(1) Active": 1, "(2) Postponed": 2 }
    ```
    ```json
    "thresholds": { "0": 955000000, "0.2": 955000001 }
    ```
  - `OptionSetCollection` — Multi-select choice field with optional sentinel wrapping
    - Example – Multi-select choice with sentinel:
      ```json
      {
        "spFieldInternalName": "Multi_x0020_Choice_x0020_Test",
        "entityAttribute": "custom_multichoicetest",
        "type": "OptionSetCollection",
        "includeSentinel": true,
        "choiceMap": {
          "Negligible": 955000000,
          "Minor": 955000001,
          "Moderate": 955000002,
          "Major": 955000003,
          "Severe": 955000004
        }
      }
      ```
      Output format in Data.xml: `[ -1,955000000,955000002,955000003,-1 ]`
  - `Boolean` or `Bool` — Exports literal `"true"`/`"false"`; optionally map with `trueValue`/`falseValue`
    - Example – Boolean field mapping:
      ```json
      {
        "spFieldInternalName": "Bool_x0020_Test",
        "entityAttribute": "custom_booltest",
        "type": "Bool"
      }
      ```
  - `Lookup` — Converts a SharePoint lookup field to a D365 lookup with entity reference
    - **⚠️ CRITICAL:** The `lookupEntity` property must be the **internal name** (logical name) of the target Dynamics 365 entity, not the display name
    - **How to find the internal entity name:**
      - Query Dataverse Web API: `GET https://yourorg.crm{region}.dynamics.com/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName&$filter=contains(DisplayName,'%27Entity%20Display%20Name%27')`
      - Check `data_schema.xml` for `<entity name="internalname" />`
      - Ask your D365 administrator for entity logical names
    - **Examples:**
      - People/User field mapped to system user table:
        ```json
        {
          "spFieldInternalName": "AssignedTo",
          "entityAttribute": "sensei_assignedto",
          "type": "Lookup",
          "lookupEntity": "systemuser"
        }
        ```
      - Lookup to a custom entity:
        ```json
        {
          "spFieldInternalName": "Program",
          "entityAttribute": "sensei_programid",
          "type": "Lookup",
          "lookupEntity": "sensei_program"
        }
        ```
  - **Image columns** — Modern SharePoint Image columns serialize values like `{ "fileName": "Reserved_ImageAttachment...png", "originalImageName": "AttachToWorkitem" }`; these are ignored during export and not written to Data.xml

- **Unsupported SharePoint column types (extend scripts if needed):** Hyperlink or Picture, Calculated (dependent expressions), Image (ignored by default), External Data, Managed Metadata/Taxonomy.

### `data_schema.xml` – CMT Schema (Optional Reference)

**Note:** This file is **optional**. The Altus DataLoader no longer requires it and validates against your live D365 environment instead.

**Recommended Use:** As a reference tool for identifying D365 entity and field names when configuring `export.config.json`.

**To create for reference:**

1. Export from your D365 environment using the **Configuration Migration Tool** (CMT)
2. Update field names and entity names to match your customizations
3. Place the updated file in this directory

**Structure:**

```xml
<?xml version="1.0"?>
<entities>
  <entity name="sensei_risk">
    <fields>
      <field name="sensei_riskid" />
      <field name="sensei_name" />
      <field name="statuscode" />
      <field name="sensei_probability" />
    </fields>
  </entity>
</entities>
```

---

## Advanced Usage

### Understanding the SharePoint Migration Dependencies

This tool is **step 3** in the overall migration process:

- **Step 1** (Previous): Project Online extraction (`0.Project-Online-Extraction/`)
  - Exports POL data; output is available as `PolExportPath`

- **Step 2** (Previous): Project migration (`1.Project-and-Resource-Migration/`)
  - Imports projects into Altus
  - **Must complete BEFORE running this SharePoint migration**

- **Step 3** (This tool): SharePoint list and document export
  - Export Lists: Use the POLExport path from step 1 (optional, for context)
  - Import Lists: **Requires projects to exist in Altus** from step 2
  - Export Documents: **File extraction only** - downloads files for use with sync tools
  - List items are linked to Altus projects during import
  - Documents are output for manual drag-into-sync or use with SharePoint sync tools

- **Step 4** (Later): Schedule migration (`2.Schedule-Migration/`)
  - Runs after projects and lists are in place

**Note on SharePoint Document Migration:**

This tool provides **document extraction only**. For complete SharePoint site migration (including lists, pages, workflows, permissions, site structure), use dedicated enterprise migration tools:

- **ShareGate** - Industry-leading SharePoint migration platform
- **Microsoft SharePoint Migration Tool** - Built-in Microsoft solution
- **Sharegate, AvePoint, or other third-party solutions** - Various specialized scenarios

These tools handle:

- ✅ Complete site structure and metadata
- ✅ Permissions and sharing settings
- ✅ Workflows and business logic
- ✅ Content types and custom fields
- ✅ Versioning and retention policies

### Document Export Modes

The document export script supports two modes:

**Simple Mode (Default):**

- Downloads files directly to library folders
- No metadata.json files created
- No version history
- Fast and efficient for basic document backup
- Perfect for creating a clean copy of documents to sync to SharePoint or other locations
- Typical use: Export documents, then drag into OneDrive sync folder or SharePoint

**Detailed Mode (`-DetailedMetadata`):**

- Creates a folder for each file
- Generates metadata.json with extended properties
- Optional version history with `-ExcludeVersionHistory:$false`
- Includes ETag, ContentType, UniqueId, and all SharePoint FieldValues
- Suitable for document backup with complete metadata preservation

**Using Exported Documents:**

1. **With OneDrive/SharePoint Sync:**

   ```
   1. Set up OneDrive sync for your new SharePoint location
   2. Drag files from Output/Documents/<Project>/ into the synced folder
   3. OneDrive automatically uploads to SharePoint
   ```

2. **For Complete SharePoint Migration:**
   ```
   Use specialized tools like ShareGate for full site migration:
   - ShareGate: Enterprise-grade SharePoint migration
   - Microsoft SharePoint Migration Tool: Native Microsoft solution
   - Other third-party tools: Various vendors support specific scenarios
   ```

### Custom Export Filters

Edit **`-Run-Migration.ps1`** to filter projects per site:

```powershell
$SiteCollections = @(
  @{
    Url           = "https://contoso.sharepoint.com/sites/pwa/"
    FolderName    = "vNext"
    ProjectFilter = @("*2024*", "Infrastructure*")  # Only 2024 projects and Infrastructure
  }
)
```

### Using Individual Scripts

For advanced scenarios, call `01-export.ps1` and `02-import.ps1` directly:

```powershell
# Custom export with specific parameters
.\1-export-lists.ps1 `
  -SiteCollectionUrl "https://tenant.sharepoint.com/sites/custom-pwa" `
  -OutputFolder ".\Output\CustomExport" `
  -ProjectFilter @("ProjectA", "ProjectB") `
  -POLExportPath "C:\exports\POL"

# Custom import with different D365 environment
.\2-import-lists.ps1 `
  -D365Url "https://custom-org.crm.dynamics.com" `
  -DataFolder ".\Output\CustomExport" `
  -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe" `
  -Force $true `
  -POLExportPath "C:\exports\POL"

# Custom document export with detailed metadata
.\3-export-documents.ps1 `
  -SiteCollectionUrl "https://tenant.sharepoint.com/sites/custom-pwa" `
  -OutputFolder ".\Output\Documents\Custom" `
  -DetailedMetadata `
  -ExcludeVersionHistory:$false `
  -ProjectFilter @("*Important*") `
  -POLExportPath "C:\exports\POL"
```

### Updating Field Mappings

To add or modify field mappings:

### Verifying Entity and Field Names

**If you have `data_schema.xml` (recommended for reference):**

1. Open the schema file
2. Find the entity you're mapping to (e.g., `<entity name="altus_risk">`)
3. Verify the field exists: `<field name="altus_probability" />`
4. Use the exact field name in your `export.config.json`

**If you don't have `data_schema.xml`:**

1. Query the Dynamics 365 Web API directly:
   ```
   GET https://yourorg.crm.dynamics.com/api/data/v9.2/EntityDefinitions(LogicalName='altus_risk')/Attributes
   ```
2. Find the field's LogicalName in the JSON response
3. Use this exact name in your `export.config.json`

**Note:** The Altus DataLoader validates against your live D365 environment, so field name errors will be caught during import.

---

## Output Structure

**List Export:**

```
Output/
├── vNext/
│   ├── Project A/
│   │   └── Data.xml
│   ├── Project B/
│   │   └── Data.xml
│   └── Export_<timestamp>.log
├── Migration1/
│   ├── Project C/
│   │   └── Data.xml
│   └── Export_<timestamp>.log
└── migration-summary.json
```

**Document Export (Simple Mode):**

```
Output/Documents/
├── vNext/
│   ├── Project A/
│   │   ├── Documents/
│   │   │   ├── file1.docx
│   │   │   ├── file2.pdf
│   │   │   ├── manifest.csv
│   │   │   └── manifest.json
│   │   └── Shared Documents/
│   │       └── ...
│   └── Project B/
│       └── ...
├── documents-manifest.csv
└── documents-manifest.json
```

**Document Export (Detailed Mode with -DetailedMetadata):**

```
Output/Documents/
├── vNext/
│   ├── Project A/
│   │   └── Documents/
│   │       ├── file1.docx/
│   │       │   ├── file1.docx
│   │       │   ├── file1_v1.0.docx  (if version history enabled)
│   │       │   └── metadata.json
│   │       ├── manifest.csv
│   │       └── manifest.json
├── documents-manifest.csv
└── documents-manifest.json
```

**migration-summary.json** contains:

- Site collection name and URL
- Operation performed (Export Lists/Import Lists/Export Documents)
- Success/Failure status
- Duration and output folder paths

---

## Troubleshooting

### "ChoiceMap for field 'X' missing value 'Y'"

**Cause:** Field value in SharePoint doesn't match any key in choiceMap.
**Solution:** Add the missing value to `export.config.json`:

```json
"choiceMap": {
  "Existing Value": 1,
  "Missing Value": 2
}
```

### "Field 'X' not found in schema" or Import Validation Errors

**Cause:** Field in `export.config.json` doesn't exist in your D365 environment with that exact name.
**Solution:**

1. Verify the field name using the Dataverse Web API:
   ```
   GET https://yourorg.crm.dynamics.com/api/data/v9.2/EntityDefinitions(LogicalName='entityname')/Attributes
   ```
2. Or use Configuration Migration Tool to export a fresh `data_schema.xml` for reference
3. Update the field name in `export.config.json` to match exactly (case-sensitive)

### Import fails with "Record exists"

**Cause:** Records already exist in D365.
**Solution:** Set `$ImportForce = $true` in `-Run-Migration.ps1` to update existing records.

### Document export taking too long

**Cause:** Detailed metadata and version history add significant processing time.
**Solution:** Use simple mode (default) for basic document backup. Only use `-DetailedMetadata` when you need extended properties, and `-ExcludeVersionHistory:$false` only when version history is required.

### "Access denied" during document export

**Cause:** User account lacks Site Collection Administrator permissions.
**Solution:** Ensure your account has Site Collection Admin role on the SharePoint site collection.

### Documents downloading but no metadata.json created

**Cause:** Running in simple mode (default behavior).
**Solution:** This is expected. Simple mode only downloads files. Use `-DetailedMetadata` if you need metadata.json files.

### Large folders/lists with many items missing some data

**Cause:** SharePoint returns items in pages. If a folder or list contains more items than SharePoint's default page size (typically 5000+), only the first page will be retrieved by the current implementation.
**Solution:** This is a known limitation affecting both list exports and document exports. For very large document libraries or lists, the scripts will need to be modified to implement explicit pagination handling. Contact your development team if you regularly encounter folders or lists exceeding the page limit.

### How do I upload the exported documents to SharePoint?

**Option 1: OneDrive Sync (Recommended)**

1. Set up OneDrive sync on your local machine for the target SharePoint folder
2. Drag files from `Output/Documents/<Project>/` into the synced folder
3. OneDrive automatically uploads them to SharePoint
4. This preserves created/modified dates and file metadata

**Option 2: Manual SharePoint Upload**

1. Open SharePoint in your browser
2. Navigate to the target document library
3. Click "Upload" and select files from `Output/Documents/<Project>/`
4. Files upload directly to SharePoint

**Option 3: For Enterprise SharePoint Migration**
Consider using **ShareGate** or other enterprise migration tools for:

- Complex site structure preservation
- Permission and sharing settings
- Content type and metadata preservation
- Workflow migration
- Retention policies and compliance settings

---

## Need Help?

- **Export issues?** Check `Output/Export_<timestamp>.log` or `Logs/DataMigration-<timestamp>.txt` for detailed error messages
- **Field mapping?** Review `export.config.json` structure and compare with SharePoint field names
- **D365 connectivity?** Reproduce connectivity issues with the approved working tool and bring back a redacted error
- **Document export?** Check manifest files for summary of what was exported
- **Performance?** Use simple mode for documents, enable detailed metadata only when needed

---

## License

For internal use only.
