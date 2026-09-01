<#
.SYNOPSIS
    Maps a named custom field from JSON to a single explicit local field in an already open MPP.

.DESCRIPTION
    For an already open Microsoft Project file, this function maps a custom field value from 
    the corresponding reporting JSON to a specified local target field.

.PARAMETER customFieldName
    The semantic name to find in JSON (e.g., "Health").

.PARAMETER localFieldTarget
    The exact local field to populate (e.g., "Text5", "Number3", "Date2", "Cost1", "Flag10").
    Only local fields in supported families (Text, Number, Date, Cost, Flag) are accepted.

.EXAMPLE
    Set-CustomFieldMapping -customFieldName "Health" -localFieldTarget "Text5"

.NOTES
    - The MPP file must already be open in Microsoft Project
    - JSON discovery follows convention: Project_*_published.mpp -> Project_*_reporting_Tasks.json
    - Task matching: TaskClientUniqueId (JSON) = UniqueID (MPP)
    - When duplicate custom field names exist, defaults to local value
    - Supports LOCAL fields only: Text, Number, Date, Cost, Flag
    - Culture-safe parsing for Number/Cost (invariant culture)
    - ISO 8601 and common date formats supported
#>

# [CmdletBinding()]
# param(
#     [Parameter(Mandatory = $true)]
#     [string]$customFieldName,

#     [Parameter(Mandatory = $true)]
#     [ValidatePattern("^(Text|Number|Date|Cost|Flag)\d+$")]
#     [string]$localFieldTarget
# )

# === LOAD COMMON HELPERS (Open-WinProj, Close-WinProj, etc.) ===
$commonPath = Join-Path $PSScriptRoot "Common.ps1"   # Adjust path if your Common.ps1 is in a different location
if (-not (Test-Path $commonPath)) {
    # Fallback: try parent folder (in case wrapper is in a subfolder)
    $commonPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.ps1"
}

if (Test-Path $commonPath) {
    . $commonPath
    Write-Verbose "Loaded common helpers from $commonPath"
} else {
    Write-Warning "Common.ps1 not found. Open-WinProj will not be available."
}

function Set-CustomFieldMapping {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$customFieldName,

        [Parameter(Mandatory = $true)]
        [ValidateScript({
                # Define valid ranges for each field family
                # Note: Text27-30 and Date10 are reserved and excluded
                $validRanges = @{
                    'Text'   = 1..26
                    'Number' = 1..20
                    'Date'   = 1..9
                    'Cost'   = 1..10
                    'Flag'   = 1..20
                }
            
                if ($_ -match '^(Text|Number|Date|Cost|Flag)(\d+)$') {
                    $family = $matches[1]
                    $index = [int]$matches[2]
                    if ($index -in $validRanges[$family]) {
                        return $true
                    }
                    throw "Invalid index '$index' for field family '$family'. Valid range: $($validRanges[$family][0])-$($validRanges[$family][-1])."
                }
                throw "Invalid localFieldTarget '$_'. Must be a local field: Text1-Text26, Number1-Number20, Date1-Date9, Cost1-Cost10, Flag1-Flag20."
            })]
        [string]$localFieldTarget
    )

    begin {
        # Extract field family (already validated by ValidateScript)
        $fieldFamily = if ($localFieldTarget -match '^(Text|Number|Date|Cost|Flag)') { $matches[1] } else { $null }

        # Initialize counters and logging
        $script:successCount = 0
        $script:skipCount = 0
        $script:errorCount = 0
        $script:startTime = Get-Date

        Write-Host "Starting custom field mapping: '$customFieldName' -> '$localFieldTarget'" -ForegroundColor Cyan
        Write-Host ("=" * 80) -ForegroundColor Cyan
        Write-Verbose "Field family: $fieldFamily (local only)"
    }

    process {
        try {
            # Step 1: Get the active Project application
            Write-Verbose "Connecting to Microsoft Project application..."
            
            $project = Connect-ToWinProj

            if (-not $project.ActiveProject) {
                Write-Error "No active project is open in Microsoft Project.`nPlease open your target MPP file in the Project window and run the mapping again."
                return
            }

            $activeProject = $project.ActiveProject
            $mppPath = $activeProject.FullName
            Write-Host "Active MPP: $mppPath" -ForegroundColor Green

            # Step 2: Discover JSON file by convention
            $jsonPath = Get-JsonPathFromMppPath -mppPath $mppPath
            if (-not $jsonPath) {
                Write-Error "Could not determine JSON path from MPP: $mppPath"
                return
            }

            if (-not (Test-Path $jsonPath)) {
                Write-Error "JSON file not found: $jsonPath"
                return
            }

            Write-Host "JSON file: $jsonPath" -ForegroundColor Green

            # Step 3: Load JSON data
            Write-Verbose "Loading JSON data..."
            try {
                $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction Stop
                $jsonData = $jsonContent | ConvertFrom-Json -ErrorAction Stop
            } catch {
                Write-Error "Failed to load or parse JSON file: $_"
                return
            }

            # Handle both formats: direct Tasks array or nested under ReportingProjectTasksData
            $tasks = $null
            if ($jsonData.ReportingProjectTasksData -and $jsonData.ReportingProjectTasksData.Tasks) {
                $tasks = $jsonData.ReportingProjectTasksData.Tasks
                Write-Verbose "Using ReportingProjectTasksData.Tasks format"
            } elseif ($jsonData.Tasks) {
                $tasks = $jsonData.Tasks
                Write-Verbose "Using direct Tasks format"
            } else {
                Write-Error "JSON file does not contain 'Tasks' array (checked both ReportingProjectTasksData.Tasks and Tasks)."
                return
            }

            Write-Host "Loaded $($tasks.Count) tasks from JSON" -ForegroundColor Green

            # Step 4: Process each task in the Project
            Write-Host "`nProcessing tasks..." -ForegroundColor Cyan
            $totalTasks = $activeProject.Tasks.Count
            $processedCount = 0

            foreach ($mppTask in $activeProject.Tasks) {
                if (-not $mppTask) { continue }  # Skip null tasks (deleted tasks)

                $processedCount++
                $taskUniqueId = $mppTask.UniqueID
                $taskName = $mppTask.Name

                Write-Verbose "[$processedCount/$totalTasks] Processing task UniqueID: $taskUniqueId, Name: $taskName"

                # Step 5: Find matching JSON task
                $jsonTask = $tasks | Where-Object { $_.TaskClientUniqueId -eq $taskUniqueId -or $_.TaskClientUniqueId -eq $taskUniqueId.ToString() }

                if (-not $jsonTask) {
                    Write-Verbose "  SKIP: No matching JSON task for UniqueID $taskUniqueId"
                    $script:skipCount++
                    continue
                }

                # Step 6: Find custom field value in JSON
                $customFieldValue = Get-CustomFieldValue -jsonTask $jsonTask -customFieldName $customFieldName

                if ($null -eq $customFieldValue) {
                    Write-Verbose "  SKIP: Custom field '$customFieldName' not found for task '$taskName' (UniqueID: $taskUniqueId)"
                    $script:skipCount++
                    continue
                }

                # Step 7: Write value to target field
                try {
                    Set-MppTaskField -mppTask $mppTask -fieldName $localFieldTarget -value $customFieldValue
                    Write-Verbose "  SUCCESS: Set $localFieldTarget = '$customFieldValue' for task '$taskName'"
                    $script:successCount++
                } catch {
                    Write-Warning "  ERROR: Failed to set $localFieldTarget for task '$taskName' (UniqueID: $taskUniqueId): $_"
                    $script:errorCount++
                }
            }

        } catch {
            Write-Error "Unexpected error during processing: $_"
            Write-Error $_.ScriptStackTrace
            return
        }
    }

    end {
        # Final summary
        $endTime = Get-Date
        $duration = $endTime - $script:startTime

        Write-Host "`n" ("=" * 80) -ForegroundColor Cyan
        Write-Host "Mapping Summary:" -ForegroundColor Cyan
        Write-Host "  Custom Field: $customFieldName" -ForegroundColor White
        Write-Host "  Target Field: $localFieldTarget" -ForegroundColor White
        Write-Host "  Success:      $script:successCount" -ForegroundColor Green
        Write-Host "  Skipped:      $script:skipCount" -ForegroundColor Yellow
        Write-Host "  Errors:       $script:errorCount" -ForegroundColor $(if ($script:errorCount -gt 0) { "Red" } else { "White" })
        Write-Host "  Duration:     $($duration.TotalSeconds.ToString('F2')) seconds" -ForegroundColor White
        Write-Host ("=" * 80) -ForegroundColor Cyan
    }
}

<#
.SYNOPSIS
    Applies multiple custom field mappings from a CSV file to the current project.
    Loads the JSON file only ONCE and reuses it for all mappings.
#>
function Set-CustomFieldMappings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MappingFile
    )

    # Step 1: Load mappings from CSV (once)
    $mappings = Import-FieldMappingConfig -MappingFile $MappingFile
    if ($mappings.Count -eq 0) {
        Write-Warning "No valid mappings to apply."
        return
    }

    Write-Host "`nApplying $($mappings.Count) custom field mappings..." -ForegroundColor Cyan

    # Step 2: Connect to Microsoft Project (once)
    $project = Connect-ToWinProj
    if (-not $project.ActiveProject) {
        Write-Error "No active project is open in Microsoft Project."
        return
    }

    $activeProject = $project.ActiveProject
    $mppPath = $activeProject.FullName
    Write-Host "Active MPP: $mppPath" -ForegroundColor Green

    # Step 3: Discover and load JSON file ONLY ONCE
    $jsonPath = Get-JsonPathFromMppPath -mppPath $mppPath
    if (-not $jsonPath -or -not (Test-Path $jsonPath)) {
        Write-Error "JSON file not found: $jsonPath"
        return
    }

    Write-Host "JSON file: $jsonPath" -ForegroundColor Green

    try {
        $jsonContent = Get-Content -Path $jsonPath -Raw -ErrorAction Stop
        $jsonData = $jsonContent | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Error "Failed to load or parse JSON file: $_"
        return
    }

    # Handle both possible JSON structures
    $tasks = $null
    if ($jsonData.ReportingProjectTasksData -and $jsonData.ReportingProjectTasksData.Tasks) {
        $tasks = $jsonData.ReportingProjectTasksData.Tasks
    } elseif ($jsonData.Tasks) {
        $tasks = $jsonData.Tasks
    } else {
        Write-Error "JSON file does not contain a Tasks array."
        return
    }

    Write-Host "Loaded $($tasks.Count) tasks from JSON" -ForegroundColor Green

    # Step 4: Apply all mappings using the same JSON data
    $success = 0
    $errors = 0
    $start = Get-Date

    foreach ($m in $mappings) {
        try {
            Write-Host "`nMapping: $($m.CustomFieldName) => $($m.LocalFieldTarget)" -ForegroundColor Cyan
            
            # Call the existing function but pass the pre-loaded tasks
            Set-CustomFieldMappingFromTasks -customFieldName $m.CustomFieldName `
                -localFieldTarget $m.LocalFieldTarget `
                -tasks $tasks `
                -Verbose:$VerbosePreference

            $success++
        } catch {
            Write-Warning "Failed '$($m.CustomFieldName)' → $($m.LocalFieldTarget): $($_.Exception.Message)"
            $errors++
        }
    }

    $duration = (Get-Date) - $start
    Write-Host "`nBulk mapping completed: $success succeeded, $errors failed ($($duration.TotalSeconds.ToString('F2'))s)" -ForegroundColor $(if ($errors -eq 0) { "Green" } else { "Yellow" })
}
#region Helper Functions

<#
.SYNOPSIS
    Internal function: Applies one custom field mapping using pre-loaded tasks data.
    This avoids reloading the JSON for every mapping.
#>
function Set-CustomFieldMappingFromTasks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$customFieldName,

        [Parameter(Mandatory = $true)]
        [string]$localFieldTarget,

        [Parameter(Mandatory = $true)]
        $tasks
    )

    $activeProject = $global:WinProjApp.ActiveProject
    if (-not $activeProject) {
        throw "No active project found."
    }

    $successCount = 0
    $skipCount = 0
    $errorCount = 0

    Write-Verbose "Processing mapping: '$customFieldName' → '$localFieldTarget'"

    foreach ($mppTask in $activeProject.Tasks) {
        if (-not $mppTask) { continue }

        $taskUniqueId = $mppTask.UniqueID
        $taskName = $mppTask.Name

        # Find matching JSON task
        $jsonTask = $tasks | Where-Object { $_.TaskClientUniqueId -eq $taskUniqueId -or $_.TaskClientUniqueId -eq $taskUniqueId.ToString() }

        if (-not $jsonTask) {
            $skipCount++
            continue
        }

        # Get custom field value
        $customFieldValue = Get-CustomFieldValue -jsonTask $jsonTask -customFieldName $customFieldName

        if ($null -eq $customFieldValue) {
            $skipCount++
            continue
        }

        # Set the value
        try {
            Set-MppTaskField -mppTask $mppTask -fieldName $localFieldTarget -value $customFieldValue
            $successCount++
        } catch {
            Write-Warning "Failed to set $localFieldTarget for task '$taskName': $_"
            $errorCount++
        }
    }

    Write-Verbose "Mapping '$customFieldName' completed: $successCount success, $skipCount skipped, $errorCount errors"
}

<#
.SYNOPSIS
    Attaches ONLY to an already-running Microsoft Project instance.
    Works in PowerShell 7+.
    Throws clear error if it cannot attach.
    No dependency on Common.ps1 or Open-WinProj.
#>
function Connect-ToWinProj {
    [CmdletBinding()]
    param()

    Write-Verbose "Attempting to connect to Microsoft Project (PS7+ compatible)..."

    # Check if Project is running at all
    $running = [System.Diagnostics.Process]::GetProcessesByName("winproj")
    if ($running.Count -eq 0) {
        throw "No Microsoft Project process is running. Please start Microsoft Project and open your MPP file first."
    }

    # Try to create the COM object - in practice this often attaches to the running instance when only one is present
    try {
        $proj = New-Object -ComObject "MSProject.Application"
        $proj.DisplayAlerts = $false
        $global:WinProjApp = $proj

        if ($proj.ActiveProject) {
            Write-Host "Attached to running Microsoft Project instance." -ForegroundColor Green
            return $proj
        } else {
            Write-Warning "Project is running but no active project is open."
            Write-Error "No active project is open. Please open your target MPP file in the visible Project window and run the mapping again."
            return $null
        }
    } catch {
        Write-Error "Failed to connect to Microsoft Project COM object: $($_.Exception.Message)"
        throw
    }
}
<#
.SYNOPSIS
    Determines the JSON file path from the MPP file path.

.DESCRIPTION
    If the MPP file follows the pattern Project_*_published.mpp, 
    returns the path to Project_*_reporting_Tasks.json in the same folder.
#>
function Get-JsonPathFromMppPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$mppPath
    )

    $directory = [System.IO.Path]::GetDirectoryName($mppPath)
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($mppPath)

    # Pattern: Project_*_published.mpp -> Project_*_reporting_Tasks.json
    if ($fileName -match '^(Project_.+)_published$') {
        $baseName = $matches[1]
        $jsonFileName = "${baseName}_reporting_Tasks.json"
        $jsonPath = Join-Path $directory $jsonFileName
        
        Write-Verbose "Derived JSON path: $jsonPath"
        return $jsonPath
    } else {
        Write-Warning "MPP filename does not match expected pattern 'Project_*_published.mpp': $fileName"
        return $null
    }
}

<#
.SYNOPSIS
    Extracts the custom field value from a JSON task object.

.DESCRIPTION
    Searches the CustomFields array for the specified field name.
    When duplicates exist, defaults to local (non-enterprise) value.
    Extracts value from #text or well-typed representations.
#>
function Get-CustomFieldValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $jsonTask,

        [Parameter(Mandatory = $true)]
        [string]$customFieldName
    )

    if (-not $jsonTask.CustomFields) {
        Write-Verbose "    Task has no CustomFields array"
        return $null
    }

    # Find all matching custom fields - check both Name and CustomFieldName properties
    $matchingFields = @($jsonTask.CustomFields | Where-Object { 
            $_.Name -eq $customFieldName -or $_.CustomFieldName -eq $customFieldName 
        })

    if ($matchingFields.Count -eq 0) {
        Write-Verbose "    Custom field '$customFieldName' not found"
        return $null
    }

    # If duplicates exist, prefer local (non-enterprise) field
    $selectedField = $null
    if ($matchingFields.Count -gt 1) {
        Write-Verbose "    Found $($matchingFields.Count) custom fields with name '$customFieldName' (duplicates detected)"
        
        # Try to find local field (typically doesn't have "Enterprise" in type or has IsLocal flag)
        # Since JSON structure may vary, we'll default to the last one or first non-null value
        $localField = $matchingFields | Where-Object { 
            $_.IsLocal -eq $true -or 
            $_.FieldType -notmatch "Enterprise" -or
            -not $_.IsEnterprise
        } | Select-Object -First 1

        if ($localField) {
            $selectedField = $localField
            Write-Verbose "    DUPLICATE RESOLUTION: Defaulted to local field instance for '$customFieldName'"
        } else {
            # Default to first field if can't determine which is local
            $selectedField = $matchingFields[0]
            Write-Warning "    DUPLICATE RESOLUTION: Multiple '$customFieldName' fields found but could not identify local; using first instance"
        }
    } else {
        $selectedField = $matchingFields[0]
    }

    # Extract value - handle both direct properties and nested CustomFieldValue structure
    $value = $null
    
    # Try nested CustomFieldValue.#text first (standard reporting format)
    if ($selectedField.CustomFieldValue -and $selectedField.CustomFieldValue.'#text') {
        $value = $selectedField.CustomFieldValue.'#text'
    }
    # Try direct #text
    elseif ($selectedField.'#text') {
        $value = $selectedField.'#text'
    }
    # Try other common value properties
    elseif ($selectedField.Value) {
        $value = $selectedField.Value
    } elseif ($selectedField.TextValue) {
        $value = $selectedField.TextValue
    } elseif ($selectedField.NumberValue) {
        $value = $selectedField.NumberValue
    } elseif ($selectedField.DateValue) {
        $value = $selectedField.DateValue
    } elseif ($selectedField.CostValue) {
        $value = $selectedField.CostValue
    }
    # Try nested CustomFieldValue with typed properties
    elseif ($selectedField.CustomFieldValue) {
        $value = $selectedField.CustomFieldValue
    }

    Write-Verbose "    Extracted value: '$value'"
    return $value
}

<#
.SYNOPSIS
    Sets a field value on an MPP task object.

.DESCRIPTION
    Writes the value to the specified local field, handling type conversion
    for Text, Number, Date, and Cost fields (local only).
    Uses culture-safe parsing for Number/Cost and ISO 8601 for dates.
#>
function Set-MppTaskField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $mppTask,

        [Parameter(Mandatory = $true)]
        [string]$fieldName,

        [Parameter(Mandatory = $true)]
        $value
    )

    if ([string]::IsNullOrWhiteSpace($value)) {
        Write-Verbose "    Value is null or empty, skipping"
        return
    }

    # Determine field type from field name (local fields only)
    $fieldType = $null
    if ($fieldName -match '^(Text|Number|Date|Cost|Flag)\d+$') {
        $fieldType = $matches[1]
    } else {
        throw "Unsupported or non-local field: $fieldName. Only Text, Number, Date, Cost, Flag (local) are supported."
    }

    # Convert value based on field type
    $convertedValue = $null
    
    try {
        switch ($fieldType) {
            "Text" {
                # Text: write as string
                $convertedValue = [string]$value
            }
            "Number" {
                # Number: culture-safe parse using invariant culture
                if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
                    $convertedValue = [double]$value
                } else {
                    $convertedValue = [double]::Parse(
                        [string]$value, 
                        [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture
                    )
                }
            }
            "Date" {
                # Date: parse ISO 8601 and common formats
                # Strip timezone offsets to use literal date/time value from JSON
                if ($value -is [DateTime]) {
                    $convertedValue = $value
                } else {
                    $stringValue = [string]$value
                    
                    # If the string has a timezone offset (ends with +/-HH:MM or Z), strip it
                    # This ensures we use the literal date/time from the JSON without conversion
                    if ($stringValue -match '^(.+?)(Z|[+-]\d{2}:\d{2})$') {
                        $stringValue = $matches[1]
                    }
                    
                    # Now parse without timezone
                    $dateFormats = @(
                        'yyyy-MM-ddTHH:mm:ss.fff',
                        'yyyy-MM-ddTHH:mm:ss',
                        'yyyy-MM-dd',
                        'MM/dd/yyyy',
                        'dd/MM/yyyy'
                    )
                    
                    $parsed = $false
                    foreach ($format in $dateFormats) {
                        try {
                            $convertedValue = [DateTime]::ParseExact(
                                $stringValue,
                                $format,
                                [System.Globalization.CultureInfo]::InvariantCulture
                            )
                            $parsed = $true
                            break
                        } catch {
                            # Try next format
                            continue
                        }
                    }
                    
                    if (-not $parsed) {
                        # Fall back to general parsing
                        $convertedValue = [DateTime]::Parse($stringValue)
                    }
                }
            }
            "Cost" {
                # Cost: culture-safe parse using invariant culture
                # Strip currency symbols before parsing
                if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
                    $convertedValue = [double]$value
                } else {
                    # Remove currency symbols and whitespace
                    $cleanValue = [string]$value -replace '[\$,\s]', ''
                    $convertedValue = [double]::Parse(
                        $cleanValue,
                        [System.Globalization.NumberStyles]::Any,
                        [System.Globalization.CultureInfo]::InvariantCulture
                    )
                }
            }
            "Flag" {
                # Flag: boolean value - true/false, yes/no, 1/0
                if ($value -is [bool]) {
                    $convertedValue = $value
                } else {
                    $stringValue = [string]$value
                    # Handle various boolean representations
                    if ($stringValue -match '^(true|yes|1)$') {
                        $convertedValue = $true
                    } elseif ($stringValue -match '^(false|no|0)$') {
                        $convertedValue = $false
                    } else {
                        # Try parsing as boolean
                        $convertedValue = [bool]::Parse($stringValue)
                    }
                }
            }
            default {
                throw "Unsupported field type: $fieldType. Only Text, Number, Date, Cost, Flag are supported."
            }
        }
    } catch {
        throw "Failed to convert value '$value' to type $fieldType : $_"
    }

    # Set the field value using SetField method
    try {
        # Access the field using the PjField enum
        # For local fields, we can use the string name directly
        $mppTask.SetField($fieldName, $convertedValue)
    } catch {
        # Fallback: Try direct property assignment
        try {
            $mppTask.$fieldName = $convertedValue
        } catch {
            throw "Failed to set field $fieldName : $_"
        }
    }
}

<#
.SYNOPSIS
    Imports a custom field mapping configuration from a CSV file.
.DESCRIPTION
    Reads a CSV file containing custom field mappings and returns an array of PSCustomObjects.
    Skips overflow fields and validates the LocalFieldTarget format.
#>
function Import-FieldMappingConfig {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$MappingFile)

    if (-not (Test-Path $MappingFile)) { throw "Mapping file not found: $MappingFile" }

    $csv = Import-Csv -Path $MappingFile
    if ($csv.Count -eq 0) {
        Write-Warning "Mapping file is empty."
        return @()
    }

    $mappings = @()
    $overflowSkipped = 0

    foreach ($row in $csv) {
        $name = $row.CustomFieldName
        $target = $row.LocalFieldTarget

        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($target)) { continue }

        if ($target -like "OVERFLOW_*") {
            $overflowSkipped++
            Write-Verbose "Skipping overflow: $name ($target)"
            continue
        }

        if ($target -notmatch '^(Text|Number|Date|Cost|Flag)\d+$') {
            Write-Warning "Invalid LocalFieldTarget '$target' for '$name'. Skipping."
            continue
        }

        $mappings += [PSCustomObject]@{
            CustomFieldName  = $name
            LocalFieldTarget = $target
        }
    }

    if ($overflowSkipped -gt 0) {
        Write-Warning "$overflowSkipped overflow fields skipped."
    }

    Write-Host "Loaded $($mappings.Count) valid mappings from $MappingFile" -ForegroundColor Green
    return $mappings
}

#endregion Helper Functions

# Export the main function (only when imported as a module)
if ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript' -and $PSScriptRoot) {
    # Being dot-sourced or run as script - no export needed
} elseif ($null -ne $MyInvocation.MyCommand.ScriptBlock.Module) {
    # Being imported as module - export the function
    Export-ModuleMember -Function Set-CustomFieldMapping, Set-CustomFieldMappings, Import-FieldMappingConfig
}
