# 6. Calendar Exceptions

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working calendar-exception scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

Extracts calendar exceptions from Project Online `*_published.xml` exports and imports them into Altus as `sensei_calendarexception` records in Dynamics 365. The export step also **provisions the enterprise calendars themselves** (`sensei_enterprisecalendar` records), so a fresh environment doesn't need calendars created by hand first.

Two tracks are supported:

| Track                       | Source                           | D365 Lookup                                                                 |
| --------------------------- | -------------------------------- | --------------------------------------------------------------------------- |
| Enterprise calendars        | `IsBaseCalendar=1`               | `sensei_enterprisecalendar` (also provisioned as its own entity, see below) |
| Bookable resource calendars | `IsBaseCalendar=0`, person-named | `sensei_bookableresource` (must already exist — see Prerequisites)          |

---

## Prerequisites

- [PowerShell 7.4 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [Azure CLI](https://aka.ms/installazurecliwindows) — used by the import step to authenticate to your Dynamics 365 environment
- `*_published.xml` files exported from Project Online (via `0.Project-Online-Extraction/ExportAllProjects.ps1` or the Microsoft export script)
- The separately approved `Altus.DevOps.D365.DataMigration.exe` obtained from the Altus Partner Portal; it is not bundled in this repository

The account running the import must have sufficient privileges in the target Altus environment (Altus Admin User recommended).

---

## Step 1 — Export (`1-export-calendar-exceptions.ps1`)

Reads all `*_published.xml` files in `-InputFolder` as a single combined batch, extracts exceptions from `Calendar/Exceptions/Exception` nodes, and writes one CMT-compliant `Data.xml` to `-OutputFolder`.

### Basic usage

```powershell
.\1-export-calendar-exceptions.ps1 `
    -InputFolder  ".\Data\Published" `
    -OutputFolder ".\Output\CalendarExceptions"
```

### Parameters

| Parameter                   | Required | Default     | Description                                                   |
| --------------------------- | -------- | ----------- | ------------------------------------------------------------- |
| `-InputFolder`              | Yes      | —           | Folder containing `*_published.xml` files                     |
| `-OutputFolder`             | Yes      | —           | Folder to write `Data.xml` and `Export_<timestamp>.log`       |
| `-ProjectFilter`            | No       | `@()` (all) | Wildcard pattern(s) on the filename prefix, e.g. `"Project*"` |
| `-ExcludeBaseCalendars`     | No       | `$false`    | Skip Track 1 (enterprise calendar exceptions)                 |
| `-ExcludeResourceCalendars` | No       | `$false`    | Skip Track 2 (bookable resource calendar exceptions)          |

### Outputs

| File                                          | Description                                                     |
| --------------------------------------------- | --------------------------------------------------------------- |
| `<OutputFolder>/Data.xml`                     | CMT-format records ready for import                             |
| `<OutputFolder>/Export_<yyyyMMdd_HHmmss>.log` | Timestamped run log (also written to `Logs/` beside the script) |

### Notes

- Records are deduplicated by `CalendarName + FromDate + ToDate` across all input files. Re-running export and re-importing with `-Force $true` upserts rather than duplicating.
- Exceptions are read exclusively from `<Exceptions><Exception>` nodes. `WeekDays` entries are not treated as exceptions.
- Non-person `IsBaseCalendar=0` calendars are skipped with a warning — classified by whether a matching `<Resource>` (via `CalendarUID`) has a populated `EmailAddress` or `NTAccount`, not by guessing at name patterns like "Milestone"/"Fixed Price" (a real pseudo-resource can still have a `<Resource>` entry, so existence alone isn't enough).
- Blank `<Name>` values fall back to `{CalendarName} – {FromDate:yyyy-MM-dd}`.
- Dates are written as `yyyy-MM-ddT00:00:00` (DateOnly, Behavior=2).

### Enterprise calendar provisioning

Alongside the calendar-exception records, the export writes one `sensei_enterprisecalendar` entity record per distinct enterprise calendar name found across the whole batch — in its own entity block, positioned _before_ `sensei_calendarexception` in `Data.xml`, so importing creates the calendars first and the exception records' lookups resolve against them.

- **Working pattern:** each provisioned calendar's `sensei_<day>isworkday` flags and `sensei_hoursperworkday` are derived from that calendar's `WeekDays/WeekDay` weekly pattern (`DayType` 1–7 = Sunday…Saturday; `DayType=0` entries are date exceptions and are not read here).
- **`Standard` → `Default Calendar`:** MS Project's default calendar always exports as `Standard`. Configure the target `Default Calendar` identifier in a working copy; the identifier is intentionally not retained in this repository.
- **Lookup by id:** every Track 1 calendar-exception record's `sensei_enterprisecalendar` lookup sets `value` to the exact id used for that calendar's record (freshly provisioned or the known `Default Calendar` id), not just `lookupentityname`.

---

## Step 2 — Import (`2-import-calendar-exceptions.ps1`)

Thin wrapper over `Altus.DevOps.D365.DataMigration.exe`. Locates `Data.xml` directly inside `-DataFolder` (the same folder you passed as `-OutputFolder` in Step 1) and imports it. The executable is supplied separately through the Altus Partner Portal.

### Basic usage

```powershell
.\2-import-calendar-exceptions.ps1 `
    -D365Url    "https://orgname.crm.dynamics.com" `
    -DataFolder ".\Output\CalendarExceptions" `
    -MigrationExePath "C:\Path\FromPartnerPortal\Altus.DevOps.D365.DataMigration.exe"
```

### Parameters

| Parameter                   | Required | Default  | Description                                                  |
| --------------------------- | -------- | -------- | ------------------------------------------------------------ |
| `-D365Url`                  | Yes      | —        | Dynamics 365 environment URL                                 |
| `-DataFolder`               | Yes      | —        | Folder containing `Data.xml` (export script `-OutputFolder`) |
| `-MigrationExePath`         | Yes      | —        | Path to the separately approved migration executable         |
| `-Force`                    | No       | `$true`  | Update existing records                                      |
| `-ParallelRequests`         | No       | `4`      | Number of parallel API requests                              |
| `-EnableDisablingOfPlugins` | No       | `$false` | Bypass Dynamics 365 plugins during import                    |

---

## Folder Structure

```
6.CalendarExceptions-Migration/
├── 1-export-calendar-exceptions.ps1   # Export — reads *_published.xml, writes Data.xml
├── 2-import-calendar-exceptions.ps1   # Import — invokes DataMigration.exe
├── Logs/                              # Run logs (script-adjacent copy)
├── Output/                            # Default output location for Data.xml and log
├── Data/Published/                    # Local diagnostic input (not tracked)
└── JPsApproach/                       # Prior exploration: COM/MPP-based resource-exception
                                        # export (Export-ProjectResourceCalendarExceptions.ps1),
                                        # kept for reference only — superseded by the two
                                        # scripts above, which read *_published.xml directly
                                        # and cover both calendar tracks in one CMT Data.xml.
```

---

## Operational Notes

- Track 2 (bookable resource) lookups require the target resource to already exist as a `sensei_bookableresource` record. If it doesn't yet, the import logs "lookup value not found" but still creates the calendar-exception record with that lookup left blank — it isn't fatal, but the record won't be linked to a resource until re-imported after the resource exists.
- The import is safe to re-run: `_recordId`s are deterministic (hashed from `CalendarName + FromDate + ToDate`, or from the calendar name for `sensei_enterprisecalendar`), and `-Force` upserts. If a run reports failures, re-running once is a reasonable first troubleshooting step before digging into `Logs/ImportErrors_<timestamp>.json` — transient SQL contention under parallel batching has been observed to clear on retry with no changes needed.
