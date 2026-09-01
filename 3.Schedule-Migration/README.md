# Altus Schedule Publish Script

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working schedule scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

## Purpose

Automates bulk publishing of Microsoft Project (.mpp) schedules to an Altus (Dataverse) environment as part of a Project Online to Altus migration (or for any already linked MPPs).

## Prerequisites

This script requires:

- Microsoft Project
- [Altus For Project](https://docs.altus.pro/products/AltusForProject/Index.html)

The script is intended to be used with its sister script `AltusPOLMigration.ps1`, however, it can be used on any linked MPPs.

The user running the script must have sufficient access to the Dataverse environment in order to read and write the required Resource and Project data. Ideally, they should be an Altus Admin User.

The user running the script must establish a connection to the target Altus environment at least once. This will ensure the authentication process operates correctly. If this has not been done, consult [here](#establish-a-connection-to-target-environment).

## Running the Script

Start Microsoft Project and ensure [Altus For Project](https://docs.altus.pro/products/AltusForProject/Index.html) Add-in is installed, active and connected to Altus.

To run the script, open PowerShell 5 (or SharePoint Online Management Shell) and navigate to the location of the script.

Select to run the following command:

```PS1
.\Publish-MPPsToAltus.ps1
```

The script can also optionally be run by passing in the following parameters:

**File Selection Parameters:**

- **All** - A switch parameter when set will tell the script to process all MPPs found within the directory. This takes precedence over NamePattern.
- **Path** - A string parameter which a user can pass to the folder to search for MPP files. If this is not set then defaults to 'Files' subfolder of the script location.
- **NamePattern** - A string parameter which a user can pass to filter the MPPs to be processed. E.g., Project\_\*.mpp
  - By default, this is set to Project\_\*\_published.mpp.

**Ordering & Preview Parameters:**

- **DateFallback** - Fallback strategy when JSON date is missing or invalid:
  - 'EndOfQueue' (default): Push files without valid dates to the end of the publish queue.
  - 'UseMppWriteTime': Use the MPP file's LastWriteTime as the sort date.
- **BaselineNumber** - (0-10) The baseline number to publish. Must be specified with IsReportable to enable baseline publishing.
- **IsReportable** - (True/False) Whether the baseline should be reportable. Must be specified with BaselineNumber to enable baseline publishing.
- **DryRun** - (Switch) Preview the publish order without publishing any files.

**Custom Field Mapping Parameters:**

- **MppFolder** - The folder where your MPP files are stored for bulk operations.
- **MappingFile** - The output of New-FieldMappingFile.ps1 script, used for bulk operations.
- **CustomFieldName** - The semantic name to find in JSON (e.g., "Health"). When specified with LocalFieldTarget, custom field mapping will be applied before publishing.
- **LocalFieldTarget** - The exact local field to populate (e.g., "Text5", "Number3"). When specified with CustomFieldName, custom field mapping will be applied before publishing.

**Retry & Resilience Parameters:**

- **RetryCount** - Number of times to retry a failed file operation (default: 3).
- **DelayBetweenFiles** - Delay in seconds between processing files to allow MS Project to stabilize (default: 2 seconds).
- **RetryDelay** - Delay in seconds before retrying a failed operation (default: 5 seconds).
- **ComRestartInterval** - Number of files to process before proactively restarting COM connection. Set to 0 to disable (default: 30).

If a user requires more info about the script, the following command can be run:

```PS1
Get-Help .\Publish-MPPsToAltus.ps1
```

### Example Usage

**Basic publish with defaults:**

```powershell
.\Publish-MPPsToAltus.ps1
```

**Publish all files with increased delays for stability:**

```powershell
.\Publish-MPPsToAltus.ps1 -All -DelayBetweenFiles 5 -RetryDelay 10
```

**Publish with aggressive COM restart for large batches:**

```powershell
.\Publish-MPPsToAltus.ps1 -All -ComRestartInterval 25 -RetryCount 5
```

**Preview publish order without actually publishing:**

```powershell
.\Publish-MPPsToAltus.ps1 -DryRun
```

### Script Workflow

On actioning the script, the following will occur:

### Script Workflow

On actioning the script, the following will occur:

1. Find Microsoft Project Interop.
2. Initialize the Microsoft Project Interop.
3. Find the Altus For Project Add-In.
4. Initialize adaptive delay and retry tracking.
5. For each project discovered in the directory (ordered by JSON metadata date):
   1. **Pre-publish cleanup**: Close any stale open projects.
   2. **Currency normalization**: Set project currency to AUD (default).
   3. **Open project**.
   4. **Connect to Altus**.
   5. **Publish schedule** to Altus.
   6. **Close and save project**.
   7. **Retry logic**: On failure, retry up to RetryCount times with adaptive delays.
   8. **COM restart**: Proactively restart COM connection every N files to prevent exhaustion.
   9. **Garbage collection**: Periodic memory cleanup every 5 files and before retries.
   10. **Add-in monitoring**: Automatically re-enable Altus add-in if it becomes disabled.

### Resilience Features

The script includes comprehensive error handling and resource management:

- **Smart Retries**: Automatically retries exceptions (COM errors, RPC errors) but not logical publish failures
- **Adaptive Delays**: Dynamically adjusts delays based on success/failure patterns:
  - Decreases delays after consecutive successes (faster processing)
  - Increases delays after failures or exceptions (allows system to stabilize)
- **Garbage Collection**: Automatic .NET garbage collection:
  - Every 5 files during normal processing
  - Before each retry attempt
  - Final cleanup at completion
- **COM Connection Management**:
  - Proactive restart every N files (default 30) to prevent COM exhaustion
  - Reactive restart on RPC/COM errors with exponential backoff
  - Tracks restart count for diagnostics
- **Add-in Recovery**: Automatically detects and re-enables the Altus add-in if Microsoft Project disables it after errors
- **Graceful Failure Handling**: Saves partial results and detailed logs if script must exit early
- **Project Cleanup**: Closes stale projects before operations to prevent conflicts

### Performance Optimization

For large batches (50+ files), consider:

- Setting **ComRestartInterval** to 25-40 based on system stability
- Increasing **DelayBetweenFiles** to 5-10 seconds
- Increasing **RetryDelay** to 10-15 seconds
- Running on a local machine (not over network)
- Closing other Office applications

The terminal will inform the following to the user:

- **File Discovery**: Total files found and publish order
- **Publish Progress**: Current file being processed (X of Y)
- **Successful Publish**: Confirmation with timestamp
- **Failed Publish**: Error message with timestamp
- **Retry Attempts**: When retrying a file after failure
- **COM Restarts**: When COM connection is restarted
- **Adaptive Delays**: Changes to delay timings based on performance
- **Add-in Status**: When add-in is re-enabled after being disabled
- **Final Summary**:
  - Count of successful publishes
  - Count of failed publishes
  - Files that failed after multiple retry attempts
  - Total COM restarts performed
  - Final adaptive delay values

### Script Outputs

The script will output to a folder following the naming convention of `Logs\Publish-MPPsToAltus_yyyyMMddHHmmss\`.

It will contain the following files:

- `publish-order.txt`
  - Detailed log showing the order files will be published based on JSON metadata dates
- `publish-order.csv`
  - CSV export of publish order with metadata for analysis
- `successes.txt`
  - A log of all successful publishes indicated by a file path.
- `failures.txt`
  - A log of all failed publishes with the following:
    - Path.
    - Error.
    - Timestamp.
    - Number of attempts.
- `failures.csv`
  - A log of all failed publishes with the following:
    - Path.
    - Error.
    - Timestamp.
    - Number of attempts.

Per project:

- `[ProjectName]-publish-log.txt`
  - A log of the publish operations performed.
- `[ProjectName]-baseline-log.txt`
  - A log of the baseline publish operations performed.
- `[ProjectName]-resource-config.json`
  - A JSON of the resource configuration created using fuzzy matching used to publish resource assignments.
  - This file documents how resources were matched/published (useful for auditing).
- `[ProjectName]-field-mapping.txt`
  - Verbose log of custom field mapping operations (if field mapping enabled).
- `[ProjectName]-field-mapping-error.txt`
  - Error details if field mapping failed (if field mapping enabled).
- `[ProjectName]-errors.json`
  - If there were warnings during publish a file with the warning details will be created.
- `[ProjectName]-baseline-errors.json`
  - If there were warnings during baseline publish a file with the warning details will be created.
- `[ProjectName]-errors.txt`
  - If there was an error during publish a file with the error details will be created.
- `[ProjectName]-null-result-error.txt`
  - If the publish operation returned a null result, details will be logged here.

## CustomFieldMapping Script

### Purpose

Maps custom field values from JSON reporting data to local fields (Text, Number, Date, Cost, Flag) in Microsoft Project MPP files. The script now supports three modes:

- Single field mapping (one field at a time)
- Bulk mapping from a CSV file (field-mappings.csv)
- Batch mode: Automatically processes all non-draft MPP files in a folder

### Prerequisites

- Microsoft Project must be installed and running
- Altus For Project add-in must be installed and enabled
- Corresponding \*\_reporting_Tasks.json files must exist alongside the MPP files
- For Bulk Mode - must have run New-FieldMappingFile.ps1 and generated a mapping file.

### How It Works

The script:

1. Connects to the active Microsoft Project application
2. Discovers the corresponding JSON file based on the open MPP filename
3. Loads task data and custom field values from the JSON
4. Matches tasks by `TaskClientUniqueId` (JSON) to `UniqueID` (MPP)
5. Maps the specified custom field value to the target local field
6. Provides detailed logging and summary statistics

### Running the Script

- Single field mapping (original):

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "<CustomFieldName>" -localFieldTarget "<LocalField>"
```

- Batch Mode (process all MPP files in a folder)

```powershell
.\Set-CustomFieldMapping.ps1 -MppFolder "C:\Path\To\Your\MPP\Files" -MappingFile ".\field-mappings.csv"
```

#### Parameters

- **MappingFile**: Output of New-FieldMappingFile.ps1
- **MppFolder**: Directory where your MPP files are stored.
- **customFieldName**: The semantic name of the custom field as it appears in the JSON data (e.g., "Health", "Budget Amount", "Flag Status")
- **localFieldTarget**: The exact local field to populate. Must match pattern: `(Text|Number|Date|Cost|Flag)\d+`
  - Supported ranges:
    - Text1-Text26 (Text27-30 are reserved)
    - Number1-Number20
    - Date1-Date9 (Date10 is reserved)
    - Cost1-Cost10
    - Flag1-Flag20

### Examples

**Map all cusom fields in all MPP:**

```powershell
.\New-FieldMappingFile.ps1 -Path "C:\Path\To\Your\MPP\Files" -OutputPath ".\field-mappings.csv"
```

```powershell
.\Set-CustomFieldMapping.ps1 -MppFolder "C:\Path\To\Your\MPP\Files" -MappingFile ".\field-mappings.csv"
```

**Map a text custom field to Text5:**

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "Health" -localFieldTarget "Text5"
```

**Map a cost custom field to Cost1:**

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "Budget Amount" -localFieldTarget "Cost1"
```

**Map a numeric custom field to Number3:**

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "Risk Score" -localFieldTarget "Number3"
```

**Map a date custom field to Date2:**

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "Deadline Date" -localFieldTarget "Date2"
```

**Map a boolean custom field to Flag10:**

```powershell
.\Set-CustomFieldMapping.ps1 -customFieldName "Flag Status" -localFieldTarget "Flag10"
```

### Data Type Handling

The script automatically handles type conversion based on the target field type:

- **Text fields**: Values are stored as strings
- **Number fields**: Culture-safe parsing using invariant culture (handles decimals like "8.5")
- **Cost fields**: Strips currency symbols (e.g., "$25,000.75" → 25000.75) and uses invariant culture parsing
- **Date fields**: Supports ISO 8601 and common date formats; strips timezone information to preserve literal date/time
- **Flag fields**: Converts boolean representations ("true"/"false", "yes"/"no", "1"/"0") to boolean values

### Duplicate Custom Field Handling

When multiple custom fields share the same name (e.g., both Enterprise and Local versions), the script:

- Detects duplicates and logs a warning
- Defaults to the local (non-enterprise) field value
- Provides verbose output indicating which instance was selected

### Script Output

The script provides real-time feedback including:

- Active MPP file path
- JSON file being used
- Task processing progress
- Success/skip/error counts per task
- Final summary with counts and duration

**Example summary output:**

```
================================================================================
Mapping Summary:
  Custom Field: Budget Amount
  Target Field: Cost1
  Success:      45
  Skipped:      12
  Errors:       0
  Duration:     2.34 seconds
================================================================================
```

### Baseline Publishing

To publish baselines along with project schedules, supply both BaselineNumber and IsReportable parameters:

```powershell
.\Publish-MPPsToAltus.ps1 -BaselineNumber 0 -IsReportable $true
```

This will publish baseline 0 as reportable after each successful project publish. Both parameters must be specified together; omitting either will skip baseline publishing entirely.

---

## Establish a Connection to Target Environment

To establish a connection to an environment, perform the following steps in Microsoft Project:

1. Open an MPP.
2. Locate the Altus tab in the ribbon.
3. Click Connect to Altus.
4. Sign in.
5. Select the environment.
6. Click Next.

Once this has been done once, there is no need to do this again (unless you are setting this up on a new machine).

## Common Issues & Resolutions

| Symptom                                                          | Cause                                                       | Resolution                                                                                        |
| ---------------------------------------------------------------- | ----------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| “Not linked to Altus” error                                      | Project missing ConnectedProjectId document property        | Manually connect via ribbon first.                                                                |
| Script exits immediately                                         | Running in PowerShell 7+                                    | Use Windows PowerShell 5.                                                                         |
| Altus automation not found                                       | Add-in not loaded / disabled                                | Re-enable in COM Add-ins manager or wait for script to auto-enable.                               |
| Add-in disabled message                                          | Microsoft Project disabled add-in after error               | User will need to re-enable to add-in                                                             |
| Many failures with resource assignment                           | Missing fuzzy mapping alignment                             | Review generated resource-config JSON and adjust resources manually then re-publish.              |
| COM exceptions / RPC_E_CALL_REJECTED                             | Stale COM connection or resource exhaustion                 | Script will automatically retry and restart COM connection. Consider lowering ComRestartInterval. |
| Project remains in memory                                        | Stale instances or abrupt termination                       | Ensure clean shutdown; close other Project windows before rerun. Script will attempt cleanup.     |
| High failure rate                                                | System overloaded or network issues                         | Increase DelayBetweenFiles and RetryDelay parameters.                                             |
| Slow performance with large batches                              | COM resource exhaustion                                     | Decrease ComRestartInterval to 20-30 files.                                                       |
| Stuck on "AFP Automation Object ready calling PublishProject..." | Publish process interrupted                                 | Cancel the script by closing the terminal and try again.                                          |
| "Microsoft Project is not running"                               | Project application not started                             | Launch Microsoft Project before running script                                                    |
| "No active project found"                                        | No MPP file is open                                         | Open the target MPP file in Project                                                               |
| "JSON file not found"                                            | MPP filename doesn't match expected pattern or JSON missing | Ensure MPP follows `Project_*_published.mpp` pattern and JSON exists                              |
| "Custom field not found"                                         | Field name doesn't match JSON data                          | Verify exact field name spelling in JSON                                                          |
| "Invalid localFieldTarget"                                       | Field pattern doesn't match supported types                 | Use format like Text5, Number3, Date2, Cost1, Flag10                                              |
| Tasks skipped                                                    | No matching JSON task or custom field missing for task      | Normal behavior when tasks don't have the custom field                                            |

## Safety / Idempotency

- Re-publishing updates existing remote records; it does not create duplicate project shells if linking already established.
- Script saves each project after publish—ensure you do not have unsaved experimental changes you wish to keep separate.
- Currency is automatically set to AUD (Australian Dollar) for all projects before publishing to ensure consistency.
- Retry logic only retries exceptions/errors, not logical publish failures, preventing duplicate operations.
- COM restart operations are logged for audit purposes.

## Performance Tips

- Run locally (avoid network latency for opening .mpp).
- Close other Office applications to reduce resource contention.
- Disable real-time antivirus scanning on the batch logs folder if performance is critical (subject to IT policy).
- For large batches (50+ files), use lower ComRestartInterval (25-30) and higher delays (5-10 seconds).
- Monitor adaptive delay changes in the output—if delays keep increasing, consider starting with higher base delays.
- Use DryRun mode first to verify file ordering before committing to full publish.
- Leverage the retry and adaptive delay features rather than trying to optimize for zero delays.

## Limitations

- No parallel publishing (sequential by design for COM stability).
- Script assumes stable add-in automation surface—update script if add-in interface changes.
- Currency is hardcoded to AUD in the Set-ProjectCurrency function (can be modified in Common.ps1 if needed).
- Field mapping requires JSON metadata files to be present and follow naming conventions.

## Key Features Summary

**Automated Retry Logic** - Configurable retry attempts with exponential backoff  
**Adaptive Delays** - Dynamic adjustment based on success/failure patterns  
**Garbage Collection** - Automatic memory management to prevent COM exhaustion  
**COM Connection Management** - Proactive and reactive restart capabilities  
**Add-in Recovery** - Automatic re-enabling of disabled add-ins  
**JSON-Based Ordering** - Publishes projects in date order from metadata  
**Custom Field Mapping** - Optional field mapping from JSON to local fields  
**Currency Normalization** - Ensures consistent AUD currency across all projects  
**Comprehensive Logging** - Detailed logs for every operation and error  
**Graceful Failure Handling** - Saves partial results if script must exit early  
**DryRun Mode** - Preview publish order without making changes

## Changelog

| Version | Date       | Notes                                                                                                                                                                                 |
| ------- | ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2.0.0   | 2026-02-25 | Added retry logic, adaptive delays, garbage collection, COM restart, add-in recovery, field mapping integration, and AUD currency normalization. Enhanced error handling and logging. |
| 1.0.0   | 2025-11-12 | Initial documented revision.                                                                                                                                                          |
