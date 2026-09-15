# Project Online to Altus Migration: Pre-Cutoff Checklist

## How to Use This Checklist

Assign an owner to every item. An item is complete only when the stated evidence has been retained in the migration archive.

This document is a planning guide. The checked-in repositories are reference material only; do not run their scripts or authenticate to customer environments. Use the current signed tools and instructions from the [Altus Partner Portal](https://partners.altus.pro).

## At-a-Glance Decisions

| Activity                                       |   Before cutoff? | Why                                                                         |
| ---------------------------------------------- | ---------------: | --------------------------------------------------------------------------- |
| Confirm scope and migration decisions          |              Yes | Determines what must be extracted                                           |
| Full Microsoft Project Online user-data export |              Yes | Provides the broader source backup                                          |
| Altus migration extraction                     |              Yes | Produces files needed for later migration                                   |
| Lookup-table export                            | Yes, if required | Lookup values may not be recoverable later                                  |
| SharePoint list and document export            | Yes, if required | Project sites may no longer be accessible                                   |
| Create or configure Altus                      |               No | Can happen before or after source extraction                                |
| Import resources and projects                  |               No | Requires the Altus target                                                   |
| Publish retained MPP schedules                 |               No | Requires a supported Windows/Microsoft Project environment and Altus target |
| Import task custom fields                      |               No | Requires projects, tasks, and schedule mapping files                        |
| Import calendar exceptions                     |               No | Requires retained XML and target resources/calendars                        |
| Import SharePoint lists                        |               No | Requires corresponding Altus projects                                       |
| UAT, training, and go-live                     |               No | Occur after target migration                                                |

## 1. Confirm Scope and Decisions

- [ ] **Confirm the project inventory.** Record each project ID, name, owner, status, and whether it is included or excluded. This prevents an incomplete export caused by an accidental project filter. Evidence: approved project list and expected project count. See the [Scope and Field Mapper](03%20-%20PO%20to%20Altus%20Migration%20-%20Scope%20and%20Field%20Mapper.xlsx) and [Migration Runsheet](01%20-%20PO%20to%20Altus%20Migration%20-%20Migration%20Runsheet.docx).

- [ ] **Confirm the resource scope.** Identify named resources, generic resources, inactive resources, and resources that must remain available for historical reporting. Evidence: source resource inventory and target matching plan. See [Project and Resources V2](Project%20Online%20Data%20Migration%20-%20Project%20%26%20Resources%20V2.pdf).

- [ ] **Decide whether draft schedules, published schedules, or both are required.** Published schedules are normally needed for the Altus schedule path; draft schedules may be required for historical retention or comparison. Evidence: decision recorded in the mapping workbook. See the [Project Online extraction reference](../0.Project-Online-Extraction/Readme.md).

- [ ] **Decide how task custom fields will be migrated.** The field may be mapped through Microsoft Project local fields or imported directly into Altus. The choice affects which JSON, XML, MPP, target-field, and mapping files must be retained. See [Task Custom Fields - Microsoft Project vs Altus](Task%20Custom%20Fields%20-%20Microsoft%20Project%20vs%20Altus.pdf) and [Migration Reference](Project%20Online%20Data%20Migration%20-%20Migration%20Reference.pdf).

- [ ] **Decide which lookup tables are required.** Record whether each lookup is flat, hierarchical, single-select, multi-select, choice-based, or many-to-many. This determines the target design and whether manual handling is needed. See the [lookup-table migration reference](../1.Lookup-Table-Migration/README.md).

- [ ] **Decide how baselines, rates, calendars, and calendar exceptions will be treated.** These are separate migration concerns and should not be assumed to follow the basic project import. See [What data gets migrated to Altus](Project%20Online%20Data%20Migration%20-%20What%20data%20gets%20migrated%20to%20Altus.pdf).

- [ ] **Agree the timesheet and historical-actuals approach.** Project Online timesheet records are not simply bulk-imported by the standard process. The protected-actuals and cutover approach must be agreed before users begin recording new actuals. See [Timesheet - Historical Actuals](Timesheet%20-%20Historical%20Actuals%20one-pager.pdf) and the [Altus for Project configuration documentation](https://docs.altus.pro/products/AltusForProject/Configuration.html).

- [ ] **Identify data requiring manual treatment.** Record master schedules, workflow history, Portfolio Analyses, formula fields, descriptions, unsupported custom fields, and any other exclusions. See the [FAQ](Project%20Online%20Data%20Migration%20-%20Frequently%20Asked%20Questions.pdf).

- [ ] **Agree the SharePoint scope.** Decide whether to migrate risks, issues, custom lists, documents, metadata, and version history. The supplied document export is not a complete SharePoint site migration. See the [SharePoint migration QRG](QRG%20-%20SharePoint%20Online%20Data%20Migration.pdf).

- [ ] **Set an internal change freeze.** Freeze source changes early enough to complete the final export, validation, and archive copy before 30 September 2026.

## 2. Prepare Source Access and Working Equipment

- [ ] **Provide a Windows operator for schedule extraction.** Microsoft Project and COM-based schedule operations require an approved Windows environment. A macOS operator can coordinate and inspect the archive but cannot perform the Windows Project automation locally. See [Altus for Project deployment](https://docs.altus.pro/products/AltusForProject/Deployment.html) and the [schedule migration reference](../3.Schedule-Migration/README.md).

- [ ] **Create or confirm the Project Online export account.** Confirm the required licence and permissions before the final extraction. See [Technical Readiness for Project Online](Project%20Online%20Data%20Migration%20-%20Technical%20Readiness%20for%20Project%20Online.pdf) and [Prerequisites](Project%20Online%20Data%20Migration%20-%20Prerequisites.pdf).

- [ ] **Confirm PWA and reporting permissions.** The operator must be able to read the projects, reporting data, resources, and schedules in scope. Record the account used and the permissions granted, but do not retain credentials or tokens. See the [Project Online extraction reference](../0.Project-Online-Extraction/Readme.md).

- [ ] **Confirm SharePoint permissions and application consent.** SharePoint export normally requires site-collection administrator access and the approved migration application permissions. See [SharePoint app permissions](Project%20Online%20Data%20Migration%20-%20Altus%20POL%20Migration%20SharePoint%20app.pdf).

- [ ] **Confirm approved tool versions.** Obtain the current tools from the Partner Portal. Do not use historical commands from this repository as operational instructions.

- [ ] **Prepare secure storage.** Ensure there is enough encrypted storage for raw exports, MPP files, XML files, SharePoint files, manifests, logs, and a second backup copy.

## 3. Complete the Required Source Exports

### 3.1 Full Project Online User-Data Export

- [ ] **Run the Microsoft Project Online user-data export, or the current approved equivalent.** This is the broader source backup and should be completed independently of the narrower Altus migration extraction. See the [Microsoft Project Online user-data export](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online).

- [ ] **Retain the complete output and export options.** Record the export date, operator, source environment, user or account context, selected options, scope, and logs. The available categories depend on the Microsoft export options used.

- [ ] **Do not use the Altus extraction as the only backup.** The Altus-oriented extractor is designed to produce migration inputs and does not export every category of Project Online data. See the [Project Online extraction reference](../0.Project-Online-Extraction/Readme.md).

### 3.2 Altus Migration Extraction

For every in-scope project:

- [ ] **Export the project inventory and source identifiers.** This is used to reconcile scope and later match records in Altus.

- [ ] **Export project and reporting JSON.** These files provide project-level values, custom fields, ownership, and reporting information for the project import. See [Project and resource migration](../2.Project-and-Resource-Migration/README.md).

- [ ] **Export task, assignment, resource, baseline, and time-phased JSON.** These files support task, financial, resource, baseline, and reporting migration where supported. Missing one category may not fail the main export but can leave a later migration step incomplete.

- [ ] **Export draft and published MPP/XML files.** The MPP files are used for later schedule publishing; the XML files are used for task and calendar mapping. See [Schedule migration](../3.Schedule-Migration/README.md) and [Calendar exceptions](../6.CalendarExceptions-Migration/README.md).

Expected historical filename patterns include:

```text
Project_<ProjectName>_draft.json
Project_<ProjectName>_published.json
Project_<ProjectName>_reporting.json
Project_<ProjectName>_reporting_Tasks.json
Project_<ProjectName>_reporting_Assignments.json
Project_<ProjectName>_reporting_Resources.json
Project_<ProjectName>_reporting_Baselines.json
Project_<ProjectName>_published.mpp
Project_<ProjectName>_published.xml
```

The current Partner Portal package may use different names or structures. Follow the current package if it differs.

### 3.3 Lookup Tables

- [ ] **Export required lookup-table values and hierarchy information.** This can be done before Altus is available. The import can happen later, but target tables must exist before values are loaded. See [Lookup-table migration](../1.Lookup-Table-Migration/README.md).

- [ ] **Record values that need manual handling.** Non-text or multi-select values may not be emitted in the form required by the migration scripts. Record the source value set and intended target representation.

### 3.4 SharePoint Lists and Documents

- [ ] **Export project-site lists.** Export risks, issues, and agreed custom lists while the project sites are accessible. List import later requires the corresponding Altus projects to exist. See the [SharePoint migration reference](../4.SharePoint-Migration/Readme.md).

- [ ] **Export documents.** Export files, folders, manifests, and any selected metadata or version history. Use detailed export mode where metadata and historical versions are required.

- [ ] **Record SharePoint features outside the exporter scope.** Capture permissions, pages, workflows, content types, retention policies, navigation, site columns, and other site configuration separately if they must be preserved. The document exporter does not perform a complete SharePoint site migration.

## 4. Retain the Migration Archive

- [ ] **Keep the full Microsoft export.** This is the historical source backup and should remain unchanged.

- [ ] **Keep the raw Altus migration package.** Do not overwrite the original MPP, XML, JSON, lookup, or SharePoint exports with files modified during migration.

- [ ] **Keep a working copy for migration.** Use a separate working copy for mapping changes, target-specific configuration, or files updated during schedule publishing.

- [ ] **Keep logs and warnings.** A successful process exit does not prove that every project exported successfully. Preserve warnings, skipped records, retries, and failures.

- [ ] **Keep a file inventory.** Record the expected and actual files per project. Include file size, modification date, export date, and optionally a cryptographic hash.

- [ ] **Keep the mapping and decision record.** Retain the completed [Scope and Field Mapper](03%20-%20PO%20to%20Altus%20Migration%20-%20Scope%20and%20Field%20Mapper.xlsx), field definitions, formulas, calendars, rates, security mappings, and manual-treatment decisions.

- [ ] **Protect the archive.** The archive may contain project data and personal information. Apply access control, encryption, retention, and backup requirements appropriate to the customer.

## 5. Validate Project Online Before Cutoff

- [ ] **Reconcile project counts.** Compare the approved scope, Project Online inventory, full export, Altus extraction, and SharePoint export counts.

- [ ] **Confirm the latest published state.** Ensure each project is published and that the published schedule represents the intended final source state. This matters because later schedule publishing uses the retained published MPP.

- [ ] **Check source identifiers.** Confirm that source project IDs, resource IDs, owner IDs, and login names are retained. These values support matching and troubleshooting.

- [ ] **Check resource matching.** Identify source users who do not have a corresponding target system user or bookable resource. Missing matches can produce missing resources or unassigned tasks. See [Project and Resources V2](Project%20Online%20Data%20Migration%20-%20Project%20%26%20Resources%20V2.pdf).

- [ ] **Inventory custom fields.** Record display names, internal names, types, labels, lookup values, formulas, calculated fields, multi-select behaviour, and target mappings. See [Task Custom Fields](Task%20Custom%20Fields%20-%20Microsoft%20Project%20vs%20Altus.pdf).

- [ ] **Check field names for export hazards.** Special characters in custom-field names can cause reporting export failures, including the documented `DatabaseUndefinedError` pattern. Escalate before changing production configuration. See the [extraction troubleshooting notes](../0.Project-Online-Extraction/Readme.md).

- [ ] **Check calendars, rates, and baselines.** Confirm which calendars, exceptions, rates, and baselines are required and that the corresponding files exist.

- [ ] **Check SharePoint site coverage.** Confirm every in-scope project site was discovered and exported. Projects without associated sites may not produce SharePoint output.

- [ ] **Check timesheet and actuals decisions.** Record the final historical actuals and cutover decision before users begin entering new data in the target. See [Timesheet - Historical Actuals](Timesheet%20-%20Historical%20Actuals%20one-pager.pdf).

- [ ] **Validate archive readability.** Open representative JSON, MPP, XML, list, document, and manifest files from the archive. Confirm that the archive can be accessed by the migration operator without modifying the raw copy.

## 6. Post-Cutoff Migration Instructions

### 6.1 Prepare the Altus Target

- [ ] **Confirm the target environment, solution version, tables, fields, choices, lookup tables, calendars, project types, roles, and security.** The live Dataverse environment is authoritative; reference schema files alone are not sufficient. See [Prerequisites](Project%20Online%20Data%20Migration%20-%20Prerequisites.pdf) and [Altus for Project deployment](https://docs.altus.pro/products/AltusForProject/Deployment.html).

- [ ] **Configure field mappings and target defaults.** Target columns must exist before custom mappings are imported. See [Migration Reference](Project%20Online%20Data%20Migration%20-%20Migration%20Reference.pdf).

### 6.2 Import Lookup Values

- [ ] **Import the retained lookup values.** Run this after the target tables and relationships exist. Validate flat versus hierarchical behaviour and confirm security access. See [Lookup-table migration](../1.Lookup-Table-Migration/README.md).

### 6.3 Preview and Import Resources and Projects

- [ ] **Run the resource and project migration in What-If mode first.** Review project matches, resource matches, unmatched users, proposed creates/updates, and warnings. See [Project and resource migration](../2.Project-and-Resource-Migration/README.md).

- [ ] **Execute the import after review.** Import resources before projects where the target process requires resource matching. Preserve the logs and record any manual corrections.

### 6.4 Publish Schedules

- [ ] **Use the retained published MPP files.** Schedule publishing can occur after the Project Online cutoff if the MPP package is complete and usable.

- [ ] **Use a supported Windows environment with Microsoft Project and Altus for Project.** The add-in must be installed, enabled, connected to the target, and permitted by the target configuration. See [Altus for Project](https://docs.altus.pro/products/AltusForProject/Index.html), [deployment](https://docs.altus.pro/products/AltusForProject/Deployment.html), and [schedule migration](../3.Schedule-Migration/README.md).

- [ ] **Preview the publish order and monitor failures.** Large batches can encounter Microsoft Project dialogs, COM failures, authentication expiry, or file-specific errors. Retain the publish logs and retry only the affected files.

### 6.5 Import Task Custom Fields

- [ ] **Run task-field migration only after projects and tasks exist.** The process uses reporting task JSON and published XML/MPP mapping files to identify the corresponding Altus tasks. See [Task-field migration](../5.TaskField-Migration/README.md).

- [ ] **Validate target field types and reserved fields.** Fields must exist in Dataverse and use the correct logical/schema names. Some task fields are controlled by the schedule and must not be overwritten.

### 6.6 Import Calendar Exceptions

- [ ] **Import enterprise calendar exceptions and resource calendar exceptions after target calendars and resources exist.** The retained published XML files provide the source exception data. See [Calendar exceptions](../6.CalendarExceptions-Migration/README.md).

- [ ] **Re-run after resources exist if resource links are missing.** A calendar exception may be created without a resource lookup when the target resource is not yet available.

### 6.7 Import SharePoint Lists and Documents

- [ ] **Import SharePoint list data after the related Altus projects exist.** The list records need target project links. See [SharePoint migration](../4.SharePoint-Migration/Readme.md).

- [ ] **Move documents using the agreed approach.** The supplied exporter downloads documents; it does not recreate a complete SharePoint site. Use an approved SharePoint migration product where permissions, pages, workflows, retention, or site structure must be preserved.

### 6.8 Complete Manual Work and UAT

- [ ] **Recreate excluded or manual data.** This may include formulas, descriptions, workflow stages, Portfolio Analyses, master-schedule relationships, and unsupported SharePoint configuration.

- [ ] **Complete UAT and post-migration validation.** Check owners, resources, project types, stages, schedules, dependencies, risks, issues, registers, documents, Teams links, and reporting. See the [Post Migration Checklist](05%20-%20Post%20Migration%20Checklist%20-%20Client.docx) and [UAT and Priority Support Process](07%20-%20PO%20to%20Altus%20Migration%20-%20UAT%20And%20Priority%20Support%20Process.docx).

## 7. Limitations and Risks

- Workflow history, Project Online timesheet records, and Portfolio Analyses are not automatically migrated by the standard process. See [What data gets migrated to Altus](Project%20Online%20Data%20Migration%20-%20What%20data%20gets%20migrated%20to%20Altus.pdf).

- Master schedules are not a direct migration target. Treat their component schedules separately.

- Formula custom fields and some project descriptions require manual recreation.

- Non-text lookup data may require manual handling.

- User and resource matching is essential. Missing target users can result in missing resources or unassigned tasks.

- SharePoint document extraction does not migrate permissions, site structure, pages, workflows, retention policies, content types, or site columns.

- Task-field import depends on the matching reporting task JSON and published XML/MPP mapping files.

- Calendar-exception import depends on retained published XML and target calendars/resources.

- Duplicate names, source filters, missing published files, whitespace differences, Microsoft Project dialogs, COM errors, and authentication expiry can create partial results.

- Remigration requires an explicit cleanup and delta strategy. Do not delete or overwrite target data without following the current [Remigration Process](Project%20Online%20Data%20Migration%20-%20Remigration%20Process.pdf).

- Current Altus documentation lists a 50 MB MPP file limit. Verify current limits in the [Altus for Project documentation](https://docs.altus.pro/products/AltusForProject/Index.html).

## 8. Source Material

- [Altus Partner Portal](https://partners.altus.pro)
- [Microsoft Project Online user-data export](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online)
- [Altus for Project overview](https://docs.altus.pro/products/AltusForProject/Index.html)
- [Altus for Project deployment](https://docs.altus.pro/products/AltusForProject/Deployment.html)
- [Altus for Project configuration](https://docs.altus.pro/products/AltusForProject/Configuration.html)
- [Project Online migration reference catalogue](index.md)
- [What data gets migrated to Altus](Project%20Online%20Data%20Migration%20-%20What%20data%20gets%20migrated%20to%20Altus.pdf)
- [Frequently Asked Questions](Project%20Online%20Data%20Migration%20-%20Frequently%20Asked%20Questions.pdf)
- [Prerequisites](Project%20Online%20Data%20Migration%20-%20Prerequisites.pdf)
- [Project and Resources V2](Project%20Online%20Data%20Migration%20-%20Project%20%26%20Resources%20V2.pdf)
- [Migration Reference](Project%20Online%20Data%20Migration%20-%20Migration%20Reference.pdf)
- [SharePoint migration reference](../4.SharePoint-Migration/Readme.md)
- [Post Migration Checklist](05%20-%20Post%20Migration%20Checklist%20-%20Client.docx)
- [UAT and Priority Support Process](07%20-%20PO%20to%20Altus%20Migration%20-%20UAT%20And%20Priority%20Support%20Process.docx)
- [Remigration Process](Project%20Online%20Data%20Migration%20-%20Remigration%20Process.pdf)
