<#
.SYNOPSIS
    Scans *_reporting_Tasks.json files and generates a custom field mapping CSV.
    Asks the user individually for every xs:decimal field (Cost vs Number).
#>

param(
    [string]$Path = ".",
    [string]$OutputPath = "field-mappings.csv"
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Local field capacities
$capacity = @{
    "Text"   = 26
    "Number" = 20
    "Date"   = 9
    "Cost"   = 10
    "Flag"   = 20
}

$fieldSlots = @{
    "Text"   = 1..26 | ForEach-Object { "Text$_" }
    "Number" = 1..20 | ForEach-Object { "Number$_" }
    "Date"   = 1..9 | ForEach-Object { "Date$_" }
    "Cost"   = 1..10 | ForEach-Object { "Cost$_" }
    "Flag"   = 1..20 | ForEach-Object { "Flag$_" }
}

$typeMap = @{
    "xs:string"   = "Text"
    "xs:int"      = "Number"
    "xs:boolean"  = "Flag"
    "xs:date"     = "Date"
    "xs:dateTime" = "Date"
}

function Get-CustomFieldType {
    param([string]$xsiType, [string]$fieldName = "")

    if ($xsiType -in @("xs:decimal", "xs:double")) {
        Write-Host "`n" + ("=" * 70) -ForegroundColor Yellow
        Write-Host "Decimal field detected:" -ForegroundColor Yellow
        Write-Host "   Field Name : $fieldName" -ForegroundColor Cyan
        Write-Host "   Appears as xs:decimal in JSON." -ForegroundColor Yellow
        Write-Host "`nShould this be mapped as a **Cost** field or a **Number** field?" -ForegroundColor Yellow

        do {
            $choice = Read-Host "Enter C for Cost or N for Number"
            $choice = $choice.Trim().ToUpper()
        } while ($choice -notin @("C", "N"))

        $chosenType = if ($choice -eq "C") { "Cost" } else { "Number" }
        Write-Host "→ This field will be treated as: $chosenType" -ForegroundColor Green
        Write-Host ("=" * 70) -ForegroundColor Yellow

        return $chosenType
    }

    # Normal mapping for other types
    if ($typeMap.ContainsKey($xsiType)) {
        return $typeMap[$xsiType]
    }
    return "Text"
}

Write-Host "Scanning JSON files in: $(Resolve-Path $Path)" -ForegroundColor Cyan

$jsonFiles = Get-ChildItem -Path $Path -Filter "*_reporting_Tasks.json" -File

if ($jsonFiles.Count -eq 0) {
    Write-Warning "No *_reporting_Tasks.json files found."
    return
}

Write-Host "Found $($jsonFiles.Count) JSON files" -ForegroundColor Green

$customFields = @{}  # Key = GUID

foreach ($file in $jsonFiles) {
    try {
        $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $tasks = $json.ReportingProjectTasksData.Tasks

        foreach ($task in $tasks) {
            foreach ($cf in $task.CustomFields) {
                $guid = $cf.CustomFieldId
                if (-not $guid -or $customFields.ContainsKey($guid)) { continue }

                # Type detection
                $xsiType = "xs:string"

                if ($cf.CustomFieldValue) {
                    $prop = $cf.CustomFieldValue.PSObject.Properties['@xsi:type']
                    if ($prop -and $prop.Value) { $xsiType = $prop.Value }
                }

                if ($xsiType -eq "xs:string") {
                    $prop = $cf.PSObject.Properties['@xsi:type']
                    if ($prop -and $prop.Value) { $xsiType = $prop.Value }
                }

                if ($xsiType -eq "xs:string") {
                    if ($cf.CustomFieldValue -is [int] -or $cf.CustomFieldValue -is [double] -or 
                        $cf.CustomFieldValue -is [decimal] -or $cf.Value -is [int] -or 
                        $cf.Value -is [double]) {
                        $xsiType = "xs:decimal"
                    }
                }

                $detectedType = Get-CustomFieldType -xsiType $xsiType -fieldName $cf.CustomFieldName

                $customFields[$guid] = [PSCustomObject]@{
                    Guid = $guid
                    Name = $cf.CustomFieldName
                    Type = $detectedType
                }
            }
        }
    } catch {
        Write-Warning "Failed to process $($file.Name): $_"
    }
}

if ($customFields.Count -eq 0) {
    Write-Warning "No custom fields found."
    return
}

# Group by type
$grouped = @{}
foreach ($type in @("Text", "Number", "Date", "Cost", "Flag")) {
    $grouped[$type] = @($customFields.Values | Where-Object Type -EQ $type | Sort-Object Guid)
}

# Build mappings
$mappings = @()
$overflowCount = @{}

foreach ($type in @("Text", "Number", "Date", "Cost", "Flag")) {
    $fields = $grouped[$type]
    if ($fields.Count -eq 0) { continue }

    $slots = $fieldSlots[$type]
    $maxSlots = $capacity[$type]
    $overflowCount[$type] = 0

    for ($i = 0; $i -lt $fields.Count; $i++) {
        $target = if ($i -lt $maxSlots) { $slots[$i] } else { "OVERFLOW_$type" }
        if ($target.StartsWith("OVERFLOW")) { $overflowCount[$type]++ }

        $mappings += [PSCustomObject]@{
            CustomFieldGuid  = $fields[$i].Guid
            CustomFieldName  = $fields[$i].Name
            LocalFieldTarget = $target
        }
    }
}

# Summary
Write-Host "`nCustom Fields Discovered:" -ForegroundColor Cyan
$anyOverflow = $false

foreach ($type in @("Text", "Number", "Date", "Cost", "Flag")) {
    $count = if ($grouped[$type]) { $grouped[$type].Count } else { 0 }
    $cap = $capacity[$type]
    
    if ($count -le $cap) {
        $icon = "Good"
        $color = "Green"
    } else {
        $icon = "Warning"
        $anyOverflow = $true
        $color = "Yellow"
    }

    $countStr = $count.ToString().PadLeft(3)
    $line = "  $($type.PadRight(7)) fields: $countStr (capacity: $cap) $icon"
    Write-Host $line -ForegroundColor $color
}

$mappings | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "`nGenerated: $OutputPath ($($mappings.Count) mappings)" -ForegroundColor Green

if (-not $anyOverflow) {
    Write-Host "All custom fields fit within available local field slots." -ForegroundColor Green
}