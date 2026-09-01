<#
.SYNOPSIS
Pester tests for PS5 -> PS7 relaunch helper.
#>

$script:ProjectRoot = Split-Path -Parent $PSScriptRoot
$script:RelaunchHelperPath = Join-Path $script:ProjectRoot "Compat\Relaunch.PS5Safe.ps1"

if (-not (Test-Path $script:RelaunchHelperPath)) {
    throw "Relaunch helper not found at $script:RelaunchHelperPath"
}

. $script:RelaunchHelperPath

Describe "Relaunch Helper Unit Tests" {
    It "Convert-SwitchLikeValuesToBoolean converts IsPresent objects to bool" {
        $params = @{
            DetailedMetadata = [ordered]@{ IsPresent = $true }
            IncludeRootWeb   = [ordered]@{ IsPresent = $false }
            Name             = "Sample"
        }

        $normalized = Convert-SwitchLikeValuesToBoolean -Parameters $params

        $normalized["DetailedMetadata"] | Should Be $true
        $normalized["IncludeRootWeb"] | Should Be $false
        $normalized["Name"] | Should Be "Sample"
    }

    It "Get-PwshRelaunchCommand emits a command that rehydrates JSON and splats params" {
        $commandText = Get-PwshRelaunchCommand -ScriptPath "C:\Temp\script.ps1" -BoundParameters @{ Name = "A" }

        $commandText | Should Match "FromBase64String"
        $commandText | Should Match "ConvertFrom-Json"
        $commandText | Should Match "@params"
        $commandText | Should Match "script\.ps1"
    }
}

Describe "Relaunch Helper Integration" {
    It "preserves arrays, bools, switches, ints and strings across PS5->PS7 handoff" {
        $powershellPath = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Path
        $pwshPath = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Path

        if (-not $powershellPath -or -not $pwshPath) {
            Write-Warning "powershell.exe or pwsh.exe not available; skipping integration assertion."
            return
        }

        $testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("relaunch-test-" + [Guid]::NewGuid().ToString("N"))
        New-Item -Path $testRoot -ItemType Directory -Force | Out-Null

        try {
            $receiverPath = Join-Path $testRoot "receiver.ps1"
            $launcherPath = Join-Path $testRoot "launcher.ps1"
            $outputPath = Join-Path $testRoot "result.json"

            @"
param(
    [string]`$OutputPath,
    [string[]]`$ProjectFilter,
    [switch]`$DetailedMetadata,
    [switch]`$IncludeRootWeb,
    [int]`$ExcludeVersionHistory = 1,
    [int]`$ParallelRequests = 0,
    [string]`$Name
)

@{
    ProjectFilter         = `$ProjectFilter
    DetailedMetadata      = [bool]`$DetailedMetadata
    IncludeRootWeb        = [bool]`$IncludeRootWeb
    ExcludeVersionHistory = `$ExcludeVersionHistory
    ParallelRequests      = `$ParallelRequests
    Name                  = `$Name
} | ConvertTo-Json -Depth 5 | Set-Content -Path `$OutputPath -Encoding UTF8

exit 0
"@ | Set-Content -Path $receiverPath -Encoding UTF8

            $escapedHelper = $script:RelaunchHelperPath.Replace("'", "''")
            $escapedReceiver = $receiverPath.Replace("'", "''")

            @"
param(
    [string]`$OutputPath,
                [string[]]`$ProjectFilter = @(),
    [switch]`$DetailedMetadata,
    [switch]`$IncludeRootWeb,
                [int]`$ExcludeVersionHistory = 1,
    [int]`$ParallelRequests = 0,
    [string]`$Name
)

. '$escapedHelper'

            # Build explicit parameter set to test serialization of mixed types reliably.
            `$boundForRelaunch = @{} + `$PSBoundParameters
            `$boundForRelaunch['ProjectFilter'] = @('Proj A', 'Proj B')
            `$boundForRelaunch['DetailedMetadata'] = [ordered]@{ IsPresent = `$true }
            `$boundForRelaunch['IncludeRootWeb'] = [ordered]@{ IsPresent = `$true }
            `$boundForRelaunch['ExcludeVersionHistory'] = 0
            `$boundForRelaunch['ParallelRequests'] = 7

            Invoke-RelaunchInPwshIfNeeded -ScriptPath '$escapedReceiver' -BoundParameters `$boundForRelaunch

throw 'Expected relaunch to exit before reaching this line.'
"@ | Set-Content -Path $launcherPath -Encoding UTF8

            & $powershellPath -NoProfile -ExecutionPolicy Bypass -File $launcherPath `
                -OutputPath $outputPath `
                -Name "Smoke"

            $LASTEXITCODE | Should Be 0
            Test-Path $outputPath | Should Be $true

            $result = Get-Content -Raw -Path $outputPath | ConvertFrom-Json
            @($result.ProjectFilter).Count | Should Be 2
            $result.ProjectFilter[0] | Should Be "Proj A"
            $result.ProjectFilter[1] | Should Be "Proj B"
            $result.DetailedMetadata | Should Be $true
            $result.IncludeRootWeb | Should Be $true
            $result.ExcludeVersionHistory | Should Be $false
            $result.ParallelRequests | Should Be 7
            $result.Name | Should Be "Smoke"
        } finally {
            if (Test-Path $testRoot) {
                Remove-Item -Path $testRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
