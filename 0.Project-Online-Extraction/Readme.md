# Script overview

> **Debugging reference only:** Scripts in this repository are intentionally blocked and cannot be executed. The working export scripts and migration tools are obtained separately from the [Altus Partner Portal](https://partners.altus.pro). Command examples below describe historical behavior for diagnosis only.

Microsoft have made available a script for exporting User data from a Project Online environment. That script is designed to run in the context of a single Project Online User and will extract all Projects (and associated data) that are relevant to that User.

The script provided in this directory is a modified version of the Microsoft script and is configured such that you can run it once per environment (rather than needing to run multiple times in a per user context) to _extract data relevant to the subsequent Altus migration scripts_.

The output from either the Microsoft script, or the script in this directory can be used in the subsequent migration steps.

> NOTE
> While the project level data exported by this script is (at the time of writing) the same as the project level data exported by the Microsoft script, this script is not intended to replace the Microsoft script in terms of backing up Project Online data. It is only intended as a helpful extraction tool to extract the project level data files that are required by the subsequent migration scripts in a single run. This script will not export _all_ available data and as such is not identical to the backup provided by the Microsoft script. We still recommend using the Microsoft script for full data backup. The script in this folder is intended only for use with the subsequent migration scripts.  
> IMPORTANT
> Project Online environments configured with Microsoft Defender for Cloud Apps may be unable to connect through the reverse-proxy path. Use the approved working tool's documentation for current environment-specific guidance.

## Microsoft Export Script

To download and use the Microsoft Project Online User Export Script, see: [Microsoft export script](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online)

## Altus-provided Export Script

The Altus-provided modified export script has the same [prerequisites](https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online#prerequisites) referred to in the Microsoft documentation.

To run the Altus-provided script, open SharePoint Online Management Shell and navigate to this directory.

Substitute the PWA Url of your Project Online environment and set the OutputDirectory to the location where you wish to save the exported files.

```powershell
.\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder
```

To export only a subset of projects by name, provide one or more wildcard filters:

```powershell
.\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder -ProjectFilter @("*Example*", "Project*")
```

```powershell
.\ExportAllProjects.ps1 -Url https://contoso.sharepoint.com/sites/pwa -OutputDirectory c:\OutputFolder -ProjectOnlineOdataFilter "ProjectStartDate gt datetime'2024-09-01T00:00:00Z'"
```

Ensure that paths which include spaces are surrounded in double quotes (")

Optionally a -Region parameter can also be provided to denote the location of the PWA environment. Valid Options are "Default", "ITAR", "Germany" and "China".

## Microsoft Project

You will need to ensure that you have Microsoft Project open while running the script as this is required to save the MPP schedule files. (Some users have also reported that you may not be required to connect to the Project Online environment via Microsoft Project itself).

Any dialogs presented by Microsoft Project while saving the mpp files will need to be interacted with manually.

If the script throws COM related errors, this can sometimes be caused by an errant Microsoft Project process still running unexpectedly. If you experience this, try these things;

- If Project is open, close it and retry the script
- If Project is closed, open it and retry the script
- If Project is closed, check Task Manager and End process for any Microsoft Project tasks, then retry the script
- Try a machine reboot and retry the script

## Results

Running either script will result in files being output to the OutputDirectory including;

- A draft and published .mpp file for each project
- A series of supporting .json files for each project

The exported files should be copied to the /2.Project-and-Resource-Migration/Files folder.

## Troubleshooting

The following error can be encountered immediately after the 'Exporting reporting data for project: {Project Name}...' log in some environments:

> DatabaseUndefinedError - `<error id="50000" name="DatabaseUndefinedError" uid="{errorid}"></error>`

The root cause of this issue looks to be related to Custom Fields in the PWA Environment containing special characters which do not look to be escaped correctly in the Microsoft data export for the Project reporting data. Altus have raised a Microsoft service request to try and get this issue rectified, but in the meantime the workaround is to update the offending Custom Field(s) in the PWA Environment to no longer contain those special characters. (After export, the Custom Field names can be updated back to include those characters again if there are implications in reporting, etc).

Known special characters that will cause this issue: '
