<#
.SYNOPSIS
    Extracts calendar exceptions from Project Online *_published.xml exports into a
    single CMT-compliant Data.xml, covering both enterprise (base) calendars and
    bookable resource calendars.

.DESCRIPTION
    Reads every *_published.xml file in -InputFolder (optionally filtered by
    -ProjectFilter) as one combined batch and extracts <Calendar><Exceptions><Exception>
    entries:

      Track 1 - Enterprise Calendar Exceptions (IsBaseCalendar=1)
        Written as sensei_calendarexception records with the sensei_enterprisecalendar
        lookup populated (matched by calendar name). One sensei_enterprisecalendar
        entity record per distinct calendar name is also provisioned ahead of these,
        so importing Data.xml creates/upserts the calendars before the lookup that
        references them by name is resolved. Each provisioned calendar's working-day
        flags (sensei_<day>isworkday) and sensei_hoursperworkday are derived from the
        calendar's WeekDays/WeekDay weekly pattern (DayType 1-7; DayType=0 entries are
        date exceptions and are not read here).

      Track 2 - Bookable Resource Calendar Exceptions (IsBaseCalendar=0, person-named)
        Written as sensei_calendarexception records with the sensei_bookableresource
        lookup populated (matched by calendar name = resource full name). Non-person
        calendars (no resource with a populated EmailAddress/NTAccount) are skipped
        with a warning.

    Records are deduplicated across all input files by CalendarName + FromDate + ToDate,
    with a deterministic _recordId so re-running export and re-importing with -Force
    upserts rather than creating duplicates.

.PARAMETER InputFolder
    Folder containing *_published.xml files. All files in the folder are treated as
    one combined batch (not per-project subfolders).

.PARAMETER OutputFolder
    Folder to write Data.xml and the timestamped Export_<yyyyMMdd_HHmmss>.log to.

.PARAMETER ProjectFilter
    Optional wildcard pattern(s) matching the filename prefix of *_published.xml files
    to process. Example: "Project*"

.PARAMETER ExcludeResourceCalendars
    Skip Track 2 (bookable resource) calendar exceptions entirely.

.PARAMETER ExcludeBaseCalendars
    Skip Track 1 (enterprise) calendar exceptions entirely.

.EXAMPLE
    .\1-export-calendar-exceptions.ps1 -InputFolder ".\Data\Published" -OutputFolder ".\Output\CalendarExceptions"

.EXAMPLE
    .\1-export-calendar-exceptions.ps1 -InputFolder ".\Data\Published" -OutputFolder ".\Output\CalendarExceptions" -ProjectFilter "Project*"
#>

param(
    [Parameter(Mandatory = $true)]
    [string] $InputFolder,

    [Parameter(Mandatory = $true)]
    [string] $OutputFolder,

    [Parameter(Mandatory = $false)]
    [string[]] $ProjectFilter = @(),

    [Parameter(Mandatory = $false)]
    [switch] $ExcludeResourceCalendars,

    [Parameter(Mandatory = $false)]
    [switch] $ExcludeBaseCalendars
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# ==============================================================================
# BOOTSTRAP
# ==============================================================================

$scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import PS5-safe relaunch helper and elevate to pwsh when required.
$relaunchHelperPath = Join-Path $scriptFolder "..\4.SharePoint-Migration\Compat\Relaunch.PS5Safe.ps1"
if (-not (Test-Path $relaunchHelperPath)) {
    Write-Error "Relaunch helper script not found: $relaunchHelperPath"
    $global:LASTEXITCODE = 1
    return
}
. $relaunchHelperPath
Invoke-RelaunchInPwshIfNeeded -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters

# Import common helpers: New-CmtDataXml, New-HashGuid, Write-LogMessage/-LogWarning/-LogError, Resolve-RelativePath.
# NOTE: Add-CmtEntityRecord and Save-ExportLog are deliberately NOT used - see local
# Add-CalendarExceptionRecord below and the direct log write in the finally block.
$commonHelpersPath = Join-Path $scriptFolder "..\4.SharePoint-Migration\CommonFunctions.ps1"
if (-not (Test-Path $commonHelpersPath)) {
    Write-Error "Common helpers script not found: $commonHelpersPath"
    $global:LASTEXITCODE = 1
    return
}
. $commonHelpersPath

Set-StrictMode -Version Latest

# ==============================================================================
# LOCAL CMT RECORD WRITER
# ==============================================================================
# Reuses New-CmtDataXml (no perf issue) but not the shared Add-CmtEntityRecord,
# whose trailing $Doc.LoadXml($Doc.OuterXml) re-serializes/re-parses the entire
# document on every call - O(n^2) at this feature's expected record volumes.

function Get-OrCreateCalendarExceptionEntityNode {
    param([xml]$Doc, [string]$EntityLogicalName)
    $root = $Doc.SelectSingleNode("/entities")
    $entityNode = $root.SelectSingleNode("entity[@name='$EntityLogicalName']")
    if (-not $entityNode) {
        $entityNode = $Doc.CreateElement("entity")
        $entityNode.SetAttribute("name", $EntityLogicalName)
        $recordsNode = $Doc.CreateElement("records")
        $entityNode.AppendChild($recordsNode) | Out-Null
        $root.AppendChild($entityNode) | Out-Null
    }
    return $entityNode
}

function Add-CalendarExceptionRecord {
    param(
        [xml]$Doc,
        [string]$EntityLogicalName,
        [hashtable]$Attributes
    )
    $entityNode = Get-OrCreateCalendarExceptionEntityNode -Doc $Doc -EntityLogicalName $EntityLogicalName
    $recordsNode = $entityNode.SelectSingleNode("records")
    $recordNode = $Doc.CreateElement("record")

    if ($Attributes.ContainsKey("_recordId")) {
        $idAttr = $Doc.CreateAttribute("id")
        $idAttr.Value = $Attributes["_recordId"]
        $recordNode.Attributes.Append($idAttr) | Out-Null
        $Attributes.Remove("_recordId")
    } else {
        $idAttr = $Doc.CreateAttribute("id")
        $idAttr.Value = [System.Guid]::NewGuid().ToString()
        $recordNode.Attributes.Append($idAttr) | Out-Null
    }

    foreach ($key in $Attributes.Keys) {
        $field = $Attributes[$key]
        $fieldNode = $Doc.CreateElement("field")
        $fieldNode.SetAttribute("name", $key)

        if ($field -is [hashtable] -and $field.ContainsKey("value")) {
            $fieldNode.SetAttribute("value", $field.value)
            if ($field.ContainsKey("lookupentity")) { $fieldNode.SetAttribute("lookupentity", $field.lookupentity) }
            if ($field.ContainsKey("lookupentityname")) { $fieldNode.SetAttribute("lookupentityname", $field.lookupentityname) }
        } else {
            $fieldNode.SetAttribute("value", [string]$field)
        }
        $recordNode.AppendChild($fieldNode) | Out-Null
    }

    $recordsNode.AppendChild($recordNode) | Out-Null
}

# ==============================================================================
# STRICT-MODE-SAFE XML NAVIGATION
# ==============================================================================
# [xml] dot-notation on Microsoft Project's MSPDI export throws
# PropertyNotFoundException under Set-StrictMode when an optional child element
# (e.g. <Exceptions>) is absent - even though the file itself is well-formed.
#
# This helper walks ChildNodes/LocalName directly rather than checking
# $Parent.PSObject.Properties.Match($ChildName): XmlElement exposes its OWN tag
# name via an intrinsic .Name property, so a property-existence check for
# "Name" incorrectly reports true even when there is no <Name> child element at
# all - e.g. an <Exception> with no <Name> child resolved to the string
# "Exception" (the element's own tag name) instead of being detected as blank,
# silently defeating the blank-name fallback. Walking ChildNodes by LocalName
# only matches genuine child elements and is also namespace-agnostic (the
# default xmlns on <Project> would make SelectNodes/SelectSingleNode with a
# plain XPath string silently return zero results without a namespace
# manager). Callers must wrap the result in @(...) - PowerShell collapses an
# empty array returned from a function to $null once captured by assignment,
# so @(...) is required at the call site too.

function Get-XmlNodeArray {
    param($Parent, [string]$ChildName)
    if ($null -eq $Parent) { return @() }
    return @($Parent.ChildNodes | Where-Object { $_.LocalName -eq $ChildName })
}

function Get-FirstXmlNode {
    param($Parent, [string]$ChildName)
    $nodes = @(Get-XmlNodeArray -Parent $Parent -ChildName $ChildName)
    if ($nodes.Count -gt 0) { return $nodes[0] }
    return $null
}

# ==============================================================================
# KNOWN ENTERPRISE CALENDAR MAPPINGS
# ==============================================================================
# MS Project's default calendar always exports as "Standard". In Altus this
# already exists as the seeded "Default Calendar" enterprise calendar record,
# so "Standard" is mapped to it rather than provisioning a duplicate. Add
# further source-name -> Altus-name mappings here as they're identified.

$script:EnterpriseCalendarNameMap = @{
    'Standard' = 'Default Calendar'
}

# Known Altus record IDs for mapped calendars, keyed by the resolved (Altus)
# name. When present, the calendar-exception lookup can reference the record
# directly by id instead of relying solely on name resolution at import time,
# and Pass 0 skips provisioning a new record for it (it already exists).
$script:KnownEnterpriseCalendarIds = @{
    'Default Calendar' = 'REPLACE-WITH-TARGET-CALENDAR-ID'
}

if ($script:KnownEnterpriseCalendarIds.Values -contains 'REPLACE-WITH-TARGET-CALENDAR-ID') {
    throw 'Calendar mappings contain placeholders. Configure target record IDs in a working copy before use.'
}

function Resolve-EnterpriseCalendarName {
    param([string]$RawName)
    if ($script:EnterpriseCalendarNameMap.ContainsKey($RawName)) {
        return $script:EnterpriseCalendarNameMap[$RawName]
    }
    return $RawName
}

# ==============================================================================
# HELPERS
# ==============================================================================

function ConvertTo-DateOnlyString {
    param([string]$DateString)
    if ([string]::IsNullOrWhiteSpace($DateString)) { return $null }
    return ([datetime]$DateString).ToString("yyyy-MM-ddT00:00:00")
}

function Get-ResourceIndex {
    <#
    Builds a CalendarUID -> {Name, EmailAddress, NTAccount} map from
    Project/Resources/Resource, used to classify Track 2 calendars as
    genuine people vs. pseudo-resources (e.g. "Payment Milestone (Fixed Price)").
    #>
    param($ProjectXml)

    $index = @{}
    $projectNode = Get-FirstXmlNode -Parent $ProjectXml -ChildName 'Project'
    if (-not $projectNode) { return $index }

    $resourcesNode = Get-FirstXmlNode -Parent $projectNode -ChildName 'Resources'
    $resources = @(Get-XmlNodeArray -Parent $resourcesNode -ChildName 'Resource')

    foreach ($resource in $resources) {
        if (-not $resource) { continue }
        $calendarUidNodes = @(Get-XmlNodeArray -Parent $resource -ChildName 'CalendarUID')
        if ($calendarUidNodes.Count -eq 0) { continue }
        $calendarUid = $calendarUidNodes[0].InnerText
        if ([string]::IsNullOrWhiteSpace($calendarUid)) { continue }

        $emailNodes = @(Get-XmlNodeArray -Parent $resource -ChildName 'EmailAddress')
        $ntAccountNodes = @(Get-XmlNodeArray -Parent $resource -ChildName 'NTAccount')
        $nameNodes = @(Get-XmlNodeArray -Parent $resource -ChildName 'Name')

        $index[$calendarUid] = [pscustomobject]@{
            Name         = if ($nameNodes.Count -gt 0) { $nameNodes[0].InnerText } else { "" }
            EmailAddress = if ($emailNodes.Count -gt 0) { $emailNodes[0].InnerText } else { "" }
            NTAccount    = if ($ntAccountNodes.Count -gt 0) { $ntAccountNodes[0].InnerText } else { "" }
        }
    }

    return $index
}

function Get-CalendarWorkingPattern {
    <#
    Derives per-day working flags and hours-per-work-day from a Calendar's
    WeekDays/WeekDay entries. Only DayType 1-7 are read (the standard MSPDI
    Sun-Sat weekly pattern) - DayType=0 entries are date-specific exceptions,
    not part of the weekly pattern, and are ignored here exactly as they are
    for the calendar-exception tracks. Hours-per-work-day is the average total
    WorkingTime span (sum of ToTime-FromTime per shift) across working days.
    #>
    param($Calendar)

    $dayTypeMap = @{
        '1' = 'Sunday'; '2' = 'Monday'; '3' = 'Tuesday'; '4' = 'Wednesday'
        '5' = 'Thursday'; '6' = 'Friday'; '7' = 'Saturday'
    }

    $days = [ordered]@{
        Sunday = $null; Monday = $null; Tuesday = $null; Wednesday = $null
        Thursday = $null; Friday = $null; Saturday = $null
    }
    $workingDayHours = @()

    $weekDaysNode = Get-FirstXmlNode -Parent $Calendar -ChildName 'WeekDays'
    $weekDays = @(Get-XmlNodeArray -Parent $weekDaysNode -ChildName 'WeekDay')

    foreach ($wd in $weekDays) {
        if (-not $wd) { continue }

        $dayTypeNodes = @(Get-XmlNodeArray -Parent $wd -ChildName 'DayType')
        if ($dayTypeNodes.Count -eq 0) { continue }
        $dayType = $dayTypeNodes[0].InnerText
        if (-not $dayTypeMap.ContainsKey($dayType)) { continue }

        $dayWorkingNodes = @(Get-XmlNodeArray -Parent $wd -ChildName 'DayWorking')
        $isWorking = ($dayWorkingNodes.Count -gt 0) -and ($dayWorkingNodes[0].InnerText -eq '1')
        $days[$dayTypeMap[$dayType]] = $isWorking

        if (-not $isWorking) { continue }

        $workingTimesNode = Get-FirstXmlNode -Parent $wd -ChildName 'WorkingTimes'
        $workingTimes = @(Get-XmlNodeArray -Parent $workingTimesNode -ChildName 'WorkingTime')
        $totalHours = 0.0
        foreach ($wt in $workingTimes) {
            if (-not $wt) { continue }
            $fromNodes = @(Get-XmlNodeArray -Parent $wt -ChildName 'FromTime')
            $toNodes = @(Get-XmlNodeArray -Parent $wt -ChildName 'ToTime')
            if ($fromNodes.Count -eq 0 -or $toNodes.Count -eq 0) { continue }
            try {
                $span = ([datetime]$toNodes[0].InnerText) - ([datetime]$fromNodes[0].InnerText)
                if ($span.TotalHours -gt 0) { $totalHours += $span.TotalHours }
            } catch { }
        }
        if ($totalHours -gt 0) { $workingDayHours += $totalHours }
    }

    $hoursPerWorkDay = $null
    if ($workingDayHours.Count -gt 0) {
        $hoursPerWorkDay = [math]::Round((($workingDayHours | Measure-Object -Average).Average), 2)
    }

    return [pscustomobject]@{
        Sunday          = $days.Sunday
        Monday          = $days.Monday
        Tuesday         = $days.Tuesday
        Wednesday       = $days.Wednesday
        Thursday        = $days.Thursday
        Friday          = $days.Friday
        Saturday        = $days.Saturday
        HoursPerWorkDay = $hoursPerWorkDay
    }
}

function Get-ExceptionRecords {
    <#
    Extracts {Name, FromDateRaw, ToDateRaw} for every Exceptions/Exception child
    of the given Calendar node. Malformed entries (no TimePeriod, or a
    TimePeriod missing FromDate/ToDate) are skipped rather than failing the run.
    #>
    param($Calendar)

    $results = @()
    $exceptionsNode = Get-FirstXmlNode -Parent $Calendar -ChildName 'Exceptions'
    $exceptions = @(Get-XmlNodeArray -Parent $exceptionsNode -ChildName 'Exception')

    foreach ($exc in $exceptions) {
        if (-not $exc) { continue }

        $timePeriodNode = Get-FirstXmlNode -Parent $exc -ChildName 'TimePeriod'
        if (-not $timePeriodNode) { continue }

        $fromDateNodes = @(Get-XmlNodeArray -Parent $timePeriodNode -ChildName 'FromDate')
        $toDateNodes = @(Get-XmlNodeArray -Parent $timePeriodNode -ChildName 'ToDate')
        if ($fromDateNodes.Count -eq 0 -or $toDateNodes.Count -eq 0) { continue }

        $fromDateRaw = $fromDateNodes[0].InnerText
        $toDateRaw = $toDateNodes[0].InnerText
        if ([string]::IsNullOrWhiteSpace($fromDateRaw) -or [string]::IsNullOrWhiteSpace($toDateRaw)) { continue }

        $excNameNodes = @(Get-XmlNodeArray -Parent $exc -ChildName 'Name')
        $excName = if ($excNameNodes.Count -gt 0) { $excNameNodes[0].InnerText } else { "" }

        $results += [pscustomobject]@{
            Name        = $excName
            FromDateRaw = $fromDateRaw
            ToDateRaw   = $toDateRaw
        }
    }

    return @($results)
}

function Write-CalendarExceptionRecord {
    <#
    Dedupes by CalendarName+FromDate+ToDate against the script-scoped $seen
    hashtable and, if new, appends a sensei_calendarexception record to $doc
    with either the enterprise or resource lookup populated. Updates the
    script-scoped counters.
    #>
    param(
        [xml]$Doc,
        [string]$CalName,
        [string]$ExcName,
        [string]$FromDateRaw,
        [string]$ToDateRaw,
        [bool]$IsBase
    )

    $fromDate = ConvertTo-DateOnlyString -DateString $FromDateRaw
    $toDate = ConvertTo-DateOnlyString -DateString $ToDateRaw

    $dedupKey = "$CalName|$fromDate|$toDate"
    if ($script:seen.ContainsKey($dedupKey)) {
        $script:dupCount++
        return
    }
    $script:seen[$dedupKey] = $true

    $excName = $ExcName
    if ([string]::IsNullOrWhiteSpace($excName)) {
        $fromDateOnly = ([datetime]$FromDateRaw).ToString("yyyy-MM-dd")
        $excName = "$CalName - $fromDateOnly"
    }

    $recordId = (New-HashGuid -InputString $dedupKey).ToString()

    $attrs = @{
        sensei_name = $excName
        sensei_from = $fromDate
        sensei_to   = $toDate
        _recordId   = $recordId
    }

    if ($IsBase) {
        # $CalName is expected to already be resolved via Resolve-EnterpriseCalendarName
        # by the caller, so e.g. "Standard" is looked up here as "Default Calendar".
        # $script:enterpriseCalendarIds is populated by Pass 0 with the exact id used
        # for that calendar's sensei_enterprisecalendar record (whether freshly
        # provisioned or a known-existing one like "Default Calendar"), so the lookup
        # can resolve by id directly instead of relying solely on name resolution.
        $lookupId = if ($script:enterpriseCalendarIds.ContainsKey($CalName)) { $script:enterpriseCalendarIds[$CalName] } else { "" }
        $attrs["sensei_enterprisecalendar"] = @{
            value            = $lookupId
            lookupentity     = "sensei_enterprisecalendar"
            lookupentityname = $CalName
        }
        Add-CalendarExceptionRecord -Doc $Doc -EntityLogicalName "sensei_calendarexception" -Attributes $attrs
        $script:baseCount++
    } else {
        $attrs["sensei_bookableresource"] = @{
            value            = ""
            lookupentity     = "sensei_bookableresource"
            lookupentityname = $CalName
        }
        Add-CalendarExceptionRecord -Doc $Doc -EntityLogicalName "sensei_calendarexception" -Attributes $attrs
        $script:resourceCount++
    }
}

# ==============================================================================
# PATH RESOLUTION / VALIDATION
# ==============================================================================

$InputFolder = Resolve-RelativePath -Path $InputFolder -BasePath $scriptFolder
$OutputFolder = Resolve-RelativePath -Path $OutputFolder -BasePath $scriptFolder

if (-not (Test-Path $InputFolder -PathType Container)) {
    Write-Error "InputFolder not found: $InputFolder"
    $global:LASTEXITCODE = 1
    return
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

Write-Host "`nCalendar Exception Export Parameters:" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "InputFolder:              $InputFolder" -ForegroundColor Gray
Write-Host "OutputFolder:             $OutputFolder" -ForegroundColor Gray
if ($ProjectFilter -and $ProjectFilter.Count -gt 0) {
    Write-Host "ProjectFilter:            $($ProjectFilter -join ', ')" -ForegroundColor Gray
}
Write-Host "ExcludeResourceCalendars: $($ExcludeResourceCalendars.IsPresent)" -ForegroundColor Gray
Write-Host "ExcludeBaseCalendars:     $($ExcludeBaseCalendars.IsPresent)" -ForegroundColor Gray
Write-Host ""

$script:logContent = @()
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

Write-LogMessage "Calendar exception export started at $(Get-Date -Format o)"
Write-LogMessage "InputFolder: $InputFolder"
Write-LogMessage "OutputFolder: $OutputFolder"

$doc = New-CmtDataXml
$seen = @{}
# Resolved calendar name -> the exact id used for its sensei_enterprisecalendar
# record (freshly provisioned in Pass 0, or a known-existing id from
# $KnownEnterpriseCalendarIds). Consulted by Write-CalendarExceptionRecord so
# Track 1 exception records can reference the calendar by id, not name alone.
$enterpriseCalendarIds = @{}

$filesProcessed = 0
$enterpriseCalendarCount = 0
$mappedCalendarCount = 0
$baseCount = 0
$resourceCount = 0
$dupCount = 0
$skippedNonPersonCount = 0

try {
    $xmlFiles = @(Get-ChildItem -Path $InputFolder -Filter "*_published.xml" -File | Sort-Object Name)

    if ($ProjectFilter -and $ProjectFilter.Count -gt 0) {
        $xmlFiles = @($xmlFiles | Where-Object {
            # Filenames follow the "Project_<ProjectName>_published.xml" convention
            # (see 0.Project-Online-Extraction/ExportDraftAndPublishedAsXML.ps1) - the
            # raw BaseName always starts with "Project_", so filtering on it directly
            # would never match a project-name pattern. Strip the fixed prefix/suffix
            # first so -ProjectFilter matches the actual project name.
            $projectName = $_.BaseName -replace '^Project_', '' -replace '_published$', ''
            $matched = $false
            foreach ($pattern in $ProjectFilter) {
                if ($projectName -like "$pattern*") { $matched = $true; break }
            }
            $matched
        })
    }

    if ($xmlFiles.Count -eq 0) {
        Write-LogWarning "No *_published.xml files matched in $InputFolder"
    }

    # Parse every file once and classify its calendars up front. This lets the
    # two passes below write ALL Track 1 (enterprise) records for the whole
    # batch before ANY Track 2 (resource) record, so Data.xml groups enterprise
    # calendar exceptions first regardless of file/calendar order in the source.
    $parsedFiles = @()

    foreach ($file in $xmlFiles) {
        Write-LogMessage "Processing file: $($file.Name)"

        try {
            [xml]$proj = Get-Content -Raw -Path $file.FullName
        } catch {
            Write-LogError "Failed to parse XML: $($file.Name)" -Exception $_.Exception.Message
            continue
        }

        $filesProcessed++

        $projectNode = Get-FirstXmlNode -Parent $proj -ChildName 'Project'
        if (-not $projectNode) {
            Write-LogWarning "No <Project> root element found in $($file.Name); skipping."
            continue
        }

        $calendarsNode = Get-FirstXmlNode -Parent $projectNode -ChildName 'Calendars'
        $calendars = @(Get-XmlNodeArray -Parent $calendarsNode -ChildName 'Calendar')
        if ($calendars.Count -eq 0) {
            continue
        }

        $resourceIndex = Get-ResourceIndex -ProjectXml $proj

        $parsedFiles += [pscustomobject]@{
            FileName      = $file.Name
            Calendars     = $calendars
            ResourceIndex = $resourceIndex
        }
    }

    # ---- Pass 0: Provision sensei_enterprisecalendar entity records ----
    # Writes one record per distinct enterprise calendar name across the whole
    # batch, ahead of the sensei_calendarexception entity block below, so that
    # when Data.xml is imported the calendars are created/upserted first and
    # the calendar-exception records' sensei_enterprisecalendar lookup (which
    # resolves by name) can find them. sensei_name is the only required field
    # on this entity - see sensei_enterprisecalendar/Entity.xml.
    if (-not $ExcludeBaseCalendars) {
        $enterpriseCalendarsSeen = @{}
        foreach ($pf in $parsedFiles) {
            foreach ($cal in $pf.Calendars) {
                if (-not $cal) { continue }

                $isBaseNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'IsBaseCalendar')
                $isBase = ($isBaseNodes.Count -gt 0) -and ($isBaseNodes[0].InnerText -eq "1")
                if (-not $isBase) { continue }

                $nameNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'Name')
                $rawCalName = if ($nameNodes.Count -gt 0) { $nameNodes[0].InnerText } else { "" }
                if ([string]::IsNullOrWhiteSpace($rawCalName)) { continue }
                $calName = Resolve-EnterpriseCalendarName -RawName $rawCalName
                if ($enterpriseCalendarsSeen.ContainsKey($calName)) { continue }
                $enterpriseCalendarsSeen[$calName] = $true

                if ($script:KnownEnterpriseCalendarIds.ContainsKey($calName)) {
                    # Already exists in Altus (e.g. "Standard" -> "Default Calendar") -
                    # don't provision a duplicate; Track 1 exceptions still reference it
                    # by its known id via Write-CalendarExceptionRecord.
                    if ($rawCalName -ne $calName) {
                        Write-LogMessage "Mapping source calendar '$rawCalName' to existing Altus calendar '$calName'"
                    }
                    $enterpriseCalendarIds[$calName] = $script:KnownEnterpriseCalendarIds[$calName]
                    $mappedCalendarCount++
                    continue
                }

                $calendarRecordId = (New-HashGuid -InputString "sensei_enterprisecalendar|$calName").ToString()
                $enterpriseCalendarIds[$calName] = $calendarRecordId

                $calendarAttrs = @{
                    sensei_name = $calName
                    _recordId   = $calendarRecordId
                }

                $pattern = Get-CalendarWorkingPattern -Calendar $cal
                if ($null -ne $pattern.Sunday)    { $calendarAttrs["sensei_sundayisworkday"]    = if ($pattern.Sunday)    { "true" } else { "false" } }
                if ($null -ne $pattern.Monday)    { $calendarAttrs["sensei_mondayisworkday"]    = if ($pattern.Monday)    { "true" } else { "false" } }
                if ($null -ne $pattern.Tuesday)   { $calendarAttrs["sensei_tuesdayisworkday"]   = if ($pattern.Tuesday)   { "true" } else { "false" } }
                if ($null -ne $pattern.Wednesday) { $calendarAttrs["sensei_wednesdayisworkday"] = if ($pattern.Wednesday) { "true" } else { "false" } }
                if ($null -ne $pattern.Thursday)  { $calendarAttrs["sensei_thursdayisworkday"]  = if ($pattern.Thursday)  { "true" } else { "false" } }
                if ($null -ne $pattern.Friday)    { $calendarAttrs["sensei_fridayisworkday"]    = if ($pattern.Friday)    { "true" } else { "false" } }
                if ($null -ne $pattern.Saturday)  { $calendarAttrs["sensei_saturdayisworkday"]  = if ($pattern.Saturday)  { "true" } else { "false" } }
                if ($null -ne $pattern.HoursPerWorkDay) {
                    $calendarAttrs["sensei_hoursperworkday"] = $pattern.HoursPerWorkDay.ToString("F2", [System.Globalization.CultureInfo]::InvariantCulture)
                }

                Add-CalendarExceptionRecord -Doc $doc -EntityLogicalName "sensei_enterprisecalendar" -Attributes $calendarAttrs
                $enterpriseCalendarCount++
            }
        }
    }

    # ---- Pass 1: Track 1 - Enterprise (base) calendar exceptions, whole batch ----
    if (-not $ExcludeBaseCalendars) {
        foreach ($pf in $parsedFiles) {
            foreach ($cal in $pf.Calendars) {
                if (-not $cal) { continue }

                $isBaseNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'IsBaseCalendar')
                $isBase = ($isBaseNodes.Count -gt 0) -and ($isBaseNodes[0].InnerText -eq "1")
                if (-not $isBase) { continue }

                $nameNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'Name')
                $calName = if ($nameNodes.Count -gt 0) { $nameNodes[0].InnerText } else { "" }
                if ([string]::IsNullOrWhiteSpace($calName)) { continue }
                $calName = Resolve-EnterpriseCalendarName -RawName $calName

                $excInfos = @(Get-ExceptionRecords -Calendar $cal)
                foreach ($info in $excInfos) {
                    Write-CalendarExceptionRecord -Doc $doc -CalName $calName -ExcName $info.Name `
                        -FromDateRaw $info.FromDateRaw -ToDateRaw $info.ToDateRaw -IsBase $true
                }
            }
        }
    }

    # ---- Pass 2: Track 2 - Bookable resource calendars, whole batch ----
    if (-not $ExcludeResourceCalendars) {
        foreach ($pf in $parsedFiles) {
            foreach ($cal in $pf.Calendars) {
                if (-not $cal) { continue }

                $isBaseNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'IsBaseCalendar')
                $isBase = ($isBaseNodes.Count -gt 0) -and ($isBaseNodes[0].InnerText -eq "1")
                if ($isBase) { continue }

                $nameNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'Name')
                $calName = if ($nameNodes.Count -gt 0) { $nameNodes[0].InnerText } else { "" }
                if ([string]::IsNullOrWhiteSpace($calName)) { continue }

                $uidNodes = @(Get-XmlNodeArray -Parent $cal -ChildName 'UID')
                $calUid = if ($uidNodes.Count -gt 0) { $uidNodes[0].InnerText } else { "" }

                $resourceInfo = $null
                if (-not [string]::IsNullOrWhiteSpace($calUid) -and $pf.ResourceIndex.ContainsKey($calUid)) {
                    $resourceInfo = $pf.ResourceIndex[$calUid]
                }

                $isPerson = $resourceInfo -and (
                    -not [string]::IsNullOrWhiteSpace($resourceInfo.EmailAddress) -or
                    -not [string]::IsNullOrWhiteSpace($resourceInfo.NTAccount)
                )

                if (-not $isPerson) {
                    Write-LogWarning "Skipping non-person resource calendar '$calName' (UID $calUid) in $($pf.FileName) - no matching resource with a populated EmailAddress/NTAccount"
                    $skippedNonPersonCount++
                    continue
                }

                $excInfos = @(Get-ExceptionRecords -Calendar $cal)
                foreach ($info in $excInfos) {
                    Write-CalendarExceptionRecord -Doc $doc -CalName $calName -ExcName $info.Name `
                        -FromDateRaw $info.FromDateRaw -ToDateRaw $info.ToDateRaw -IsBase $false
                }
            }
        }
    }

    $outputDataPath = Join-Path $OutputFolder "Data.xml"
    $doc.Save($outputDataPath)

    $exceptionRecords = $baseCount + $resourceCount
    $totalRecords = $enterpriseCalendarCount + $exceptionRecords

    Write-Host ""
    Write-LogMessage "===================================="
    Write-LogMessage "Calendar Exception Export Summary"
    Write-LogMessage "===================================="
    Write-LogMessage "XML files processed:              $filesProcessed"
    Write-LogMessage "Enterprise calendars provisioned:  $enterpriseCalendarCount"
    Write-LogMessage "Enterprise calendars mapped:       $mappedCalendarCount"
    Write-LogMessage "Base-calendar exceptions:          $baseCount"
    Write-LogMessage "Resource-calendar exceptions:      $resourceCount"
    Write-LogMessage "Non-person calendars skipped:      $skippedNonPersonCount"
    Write-LogMessage "Duplicates eliminated:             $dupCount"
    Write-LogMessage "Total records written to Data.xml: $totalRecords"
    Write-LogMessage "Data.xml written to: $outputDataPath"

    $global:LASTEXITCODE = 0
}
catch {
    Write-LogError "Export failed" -Exception $_.Exception.Message
    $global:LASTEXITCODE = 1
    throw
}
finally {
    Write-LogMessage "Export finished at $(Get-Date -Format o)"
    # Intentionally NOT calling Save-ExportLog: its Logs folder resolves relative to
    # CommonFunctions.ps1's own directory (4.SharePoint-Migration), not this script's,
    # so it would write into the wrong migration folder. Write the AC-mandated log
    # directly instead.
    $logFilePath = Join-Path $OutputFolder "Export_$timestamp.log"
    $script:logContent | Out-File -FilePath $logFilePath -Encoding UTF8 -Force
    Write-Host "Log saved to: $logFilePath" -ForegroundColor Gray
}
