# FinanceMigration.ps1

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [string]$ProjectName,
    [string]$TasksFile,
    [string]$BaselinesFile,
    [string]$PublishedFile,
    [string]$Directory,
    [switch]$DryRun,
    [string]$LogFile = "finance_migration_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

. $PSScriptRoot\Core.ps1
. $PSScriptRoot\TableOperations.ps1

$logsFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory | Out-Null
}
$logFile = Join-Path $logsFolder $LogFile

function Write-LogMessage {
    param ([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

Write-LogMessage "Starting migration" "INFO"
if ($DryRun) { Write-LogMessage "DRY-RUN MODE" "WARN" }

$dataverseUrl = Read-Host "Dataverse URL"
Connect $dataverseUrl

$orgQuery = "?`$select=basecurrencyid&`$expand=basecurrencyid(`$select=isocurrencycode)&`$top=1"
$orgResult = Get-Records -setName 'organizations' -query $orgQuery

if ($orgResult.value.Count -eq 0 -or -not $orgResult.value[0].basecurrencyid) {
    Write-LogMessage "ERROR: Could not retrieve organization or base currency object. Aborting." "ERROR"
    exit 1
}

$baseCurrencyCode = $orgResult.value[0].basecurrencyid.isocurrencycode

if (-not $baseCurrencyCode) {
    Write-LogMessage "ERROR: Base currency ISO code not found in expanded object. Aborting." "ERROR"
    exit 1
}

$baseCurrencyCode = $baseCurrencyCode.Trim().ToUpper()
Write-LogMessage "Environment base currency: $baseCurrencyCode" "INFO"

$projectSets = @()

if ($Directory) {
    if (-not (Test-Path $Directory -PathType Container)) {
        Write-LogMessage "Directory not found: $Directory" "ERROR"
        exit 1
    }

    $allJson = Get-ChildItem $Directory -Filter "*.json" -File

    $groups = @{}
    foreach ($f in $allJson) {
        if ($f.Name -match '^Project_(.*?)_(reporting_Tasks|reporting_Baselines|published)\.json$') {
            $name = $matches[1] -replace '_', ' '
            if (-not $groups[$name]) { $groups[$name] = @{} }
            $suffix = $matches[2]
            $groups[$name][$suffix] = $f.FullName
        }
    }

    foreach ($name in $groups.Keys) {
        $g = $groups[$name]
        if ($g['reporting_Tasks'] -and $g['reporting_Baselines'] -and $g['published']) {
            $projectSets += $g
            Write-LogMessage "Complete: $name"
        } else {
            Write-LogMessage "Incomplete: $name" "WARN"
        }
    }
} else {
    $t = $TasksFile
    $b = $BaselinesFile
    $p = $PublishedFile

    if ($ProjectName -and -not $t -and -not $b -and -not $p) {
        $safe = $ProjectName -replace '\s+', '_'
        $t = "Project_${safe}_reporting_Tasks.json"
        $b = "Project_${safe}_reporting_Baselines.json"
        $p = "Project_${safe}_published.json"
    }

    if (-not ($t -and $b -and $p)) {
        Write-LogMessage "Missing files" "ERROR"
        exit 1
    }

    $missing = @($t, $b, $p) | Where-Object { -not (Test-Path $_) }
    if ($missing) {
        Write-LogMessage "Missing: $($missing -join ', ')" "ERROR"
        exit 1
    }

    $projectSets += @{ Tasks = $t; Baselines = $b; Published = $p }
}

if (-not $projectSets) {
    Write-LogMessage "No projects" "ERROR"
    exit 1
}

Write-LogMessage "Processing $($projectSets.Count) projects"

Invoke-DataverseCommands {
    foreach ($s in $projectSets) {
        $tasksPath = $s['reporting_Tasks']
        $baselinesPath = $s['reporting_Baselines']
        $publishedPath = $s['published']

        if (-not $tasksPath -or -not (Test-Path $tasksPath)) {
            Write-LogMessage "ERROR: Tasks file missing or invalid: $tasksPath" "ERROR"
            continue
        }
        if (-not $baselinesPath -or -not (Test-Path $baselinesPath)) {
            Write-LogMessage "ERROR: Baselines file missing or invalid: $baselinesPath" "ERROR"
            continue
        }
        if (-not $publishedPath -or -not (Test-Path $publishedPath)) {
            Write-LogMessage "ERROR: Published file missing or invalid: $publishedPath" "ERROR"
            continue
        }

        $tj = Get-Content $tasksPath -Raw | ConvertFrom-Json
        $bj = Get-Content $baselinesPath -Raw | ConvertFrom-Json
        $pj = Get-Content $publishedPath -Raw | ConvertFrom-Json

        $proj = $pj.NewDataSet.Project
        $name = $proj.ProjectName
        $start = $proj.ProjectStartDate

        $currCode = if ($proj.ProjectCurrencyCode) { $proj.ProjectCurrencyCode } else { "Unknown" }
        $currSym = if ($proj.ProjectCurrencySymbol) { $proj.ProjectCurrencySymbol } else { "" }
        $currDig = if ($proj.ProjectCurrencyDigits) { $proj.ProjectCurrencyDigits } else { "2" }

        $invariant = [System.Globalization.CultureInfo]::InvariantCulture

        try {
            $parsedStart = [DateTime]::ParseExact($start, "MM/dd/yyyy HH:mm:ss", $invariant)
        } catch {
            try {
                $parsedStart = [DateTime]::ParseExact($start, "dd/MM/yyyy HH:mm:ss", $invariant)
            } catch {
                try {
                    $parsedStart = [DateTime]::Parse($start, $invariant)
                } catch {
                    Write-LogMessage "ERROR: Cannot parse date '$start' for project '$name'. Skipping date-related fields." "ERROR"
                    continue
                }
            }
        }

        $startUtc = $parsedStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $startDateOnly = $parsedStart.ToString("yyyy-MM-dd")

        Write-LogMessage "Project: $name | Currency: $currCode ($currSym, $currDig) | Start: $startUtc"

        $jsonCurr = if ($currCode -and $currCode -match '^[A-Z]{3}$') { $currCode.Trim().ToUpper() } else { "UNKNOWN" }

        if ($jsonCurr -ne $baseCurrencyCode) {
            Write-LogMessage "Mismatched Currency, Local: $jsonCurr, Remote: $baseCurrencyCode - skipping $name" "WARN"
            continue
        }

        Write-LogMessage "Currency OK ($jsonCurr matches base). Proceeding." "INFO"

        $q = "?`$filter=sensei_name eq '$([uri]::EscapeDataString($name))'&`$select=sensei_projectid,sensei_name&`$top=10"
        $projects = Get-Records -setName 'sensei_projects' -query $q

        $projectId = $null
        if ($projects.value.Count -eq 0) {
            $manual = Read-Host "Manual projectid for '$name' (or Enter to skip)"
            if ($manual) { $projectId = $manual }
        } elseif ($projects.value.Count -eq 1) {
            $projectId = $projects.value[0].sensei_projectid
            $ok = Read-Host "Use $($projects.value[0].sensei_name) ($projectId)? (y/n)"
            if ($ok -notmatch '^[yY]$') { continue }
        } else {
            Write-LogMessage "Multiple matches for '$name':"
            $i = 1
            foreach ($p in $projects.value) {
                Write-LogMessage "  [$i] $($p.sensei_name) ($($p.sensei_projectid))"
                $i++
            }
            $sel = Read-Host "Select # (or Enter skip)"
            if ($sel -match '^\d+$' -and $sel -ge 1 -and $sel -le $projects.value.Count) {
                $projectId = $projects.value[[int]$sel - 1].sensei_projectid
            }
        }

        if (-not $projectId) { continue }

        $iq = "?`$filter=_sensei_project_value eq $projectId&`$top=1"
        $tq = "?`$filter=_sensei_project_value eq $projectId&`$top=1"

        $exI = Get-Records -setName 'sensei_financialitems' -query $iq
        $exT = Get-Records -setName 'sensei_financialtransactions' -query $tq

        if ($exI.value -or $exT.value) {
            $ok = Read-Host "Existing data found. Continue? (y/n)"
            if ($ok -notmatch '^[yY]$') { continue }
        }

        $taskTotal = $tj.ReportingProjectTasksData.Tasks.Count
        $summarySkipped = ($tj.ReportingProjectTasksData.Tasks | Where-Object { $_.TaskIsSummary -eq "true" }).Count

        Write-LogMessage "$name - Tasks: $taskTotal total, $summarySkipped summary skipped"

        $summary = @{}
        foreach ($t in $tj.ReportingProjectTasksData.Tasks) {
            $summary[$t.TaskUID] = $t.TaskIsSummary -eq "true"
        }

        $budget = 0
        foreach ($b in $bj.ReportingProjectBaselinesData.TaskBaseline) {
            if ($b.BaselineNumber -eq "0" -and -not $summary[$b.TaskUID]) {
                $budget += [decimal]$b.TaskBaselineCost
            }
        }

        $cost = 0; $actual = 0
        foreach ($t in $tj.ReportingProjectTasksData.Tasks) {
            if (-not $summary[$t.TaskUID]) {
                $cost += [decimal]$t.TaskCost
                $actual += [decimal]$t.TaskActualCost
            }
        }

        Write-LogMessage "Budget: $budget | Cost: $cost | Actual: $actual"

        $entries = @(
            @{Type = 955000000; Name = "$name Budget"; ExtId = "$projectId`_BUDGET_ITEM"; TExt = "$projectId`_BUDGET_TRANS"; Amt = $budget; TType = 955000000 },
            @{Type = 955000001; Name = "$name Cost"; ExtId = "$projectId`_COST_ITEM"; TExt = "$projectId`_COST_TRANS"; Amt = $cost; TType = 955000001 },
            @{Type = 955000001; Name = "$name Actual"; ExtId = "$projectId`_ACTUAL_ITEM"; TExt = "$projectId`_ACTUAL_TRANS"; Amt = $actual; TType = 955000002 }
        )

        $inc = 1
        foreach ($e in $entries) {
            if ($e.Amt -eq 0 -and -not $DryRun) { continue }

            $iBody = @{
                sensei_name                 = $e.Name
                sensei_type                 = $e.Type
                sensei_externalid           = $e.ExtId
                "sensei_project@odata.bind" = "/sensei_projects($projectId)"
            }

            $iF = "sensei_externalid eq '$($e.ExtId)' and _sensei_project_value eq $projectId"
            $exI = Get-Records -setName 'sensei_financialitems' -query "?`$filter=$iF&`$select=sensei_financialitemid"

            $iid = $null
            if ($exI.value) {
                $iid = $exI.value[0].sensei_financialitemid
                if (-not $DryRun) { Update-Record -setName 'sensei_financialitems' -id $iid -body $iBody }
                Write-LogMessage "Updated item $($e.Name) ($iid)"
            } else {
                if (-not $DryRun) {
                    $iid = New-Record -setName 'sensei_financialitems' -body $iBody
                } else {
                    $iid = "DRY-ITEM-$([guid]::NewGuid())"
                }
                Write-LogMessage "Created item $($e.Name) ($iid)"
            }

            $tName = "FT-{0:D4}" -f $inc++
            $tBody = @{
                sensei_name                       = $tName
                sensei_type                       = $e.TType
                sensei_value                      = $e.Amt
                sensei_date                       = $startDateOnly
                sensei_externalid                 = $e.TExt
                "sensei_project@odata.bind"       = "/sensei_projects($projectId)"
                "sensei_financialitem@odata.bind" = "/sensei_financialitems($iid)"
            }

            $tF = "sensei_externalid eq '$($e.TExt)' and _sensei_project_value eq $projectId"
            $exT = Get-Records -setName 'sensei_financialtransactions' -query "?`$filter=$tF&`$select=sensei_financialtransactionid"

            if ($exT.value) {
                $tid = $exT.value[0].sensei_financialtransactionid
                if (-not $DryRun) { Update-Record -setName 'sensei_financialtransactions' -id $tid -body $tBody }
                Write-LogMessage "Updated trans $tName ($tid)"
            } else {
                if (-not $DryRun) {
                    $tid = New-Record -setName 'sensei_financialtransactions' -body $tBody
                } else {
                    $tid = "DRY-TRANS-$([guid]::NewGuid())"
                }
                Write-LogMessage "Created trans $tName ($tid)"
            }
        }

        Write-LogMessage "Completed: $name"
    }
}

Write-LogMessage "Migration finished" "INFO"