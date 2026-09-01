<#
.SYNOPSIS
Pester coverage for the debugging-only execution boundary.
#>

$script:RepositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:PowerShellPath = (Get-Command pwsh -ErrorAction Stop).Path

function Get-GuardedEntryPoints {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProbePath
    )

    @(
        [PSCustomObject]@{
            RelativePath = "0.Project-Online-Extraction/ExportAllProjects.ps1"
            Arguments    = @("-Url", "https://example.invalid", "-OutputDirectory", $ProbePath)
        }
        [PSCustomObject]@{
            RelativePath = "0.Project-Online-Extraction/ExportProjectUserContent.ps1"
            Arguments    = @("-Url", "https://example.invalid", "-LoginName", "diagnostic@example.invalid", "-OutputDirectory", $ProbePath)
        }
        [PSCustomObject]@{
            RelativePath = "0.Project-Online-Extraction/Invoke-RedactProjectUser.ps1"
            Arguments    = @("-Url", "https://example.invalid", "-LoginName", "diagnostic@example.invalid")
        }
        [PSCustomObject]@{
            RelativePath = "1.Lookup-Table-Migration/ExportLookupTableData.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "1.Lookup-Table-Migration/ImportLookupTableData.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "2.Project-and-Resource-Migration/AltusPOLMigration.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "2.Project-and-Resource-Migration/BYOPrincipal-GrantPermission.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "2.Project-and-Resource-Migration/FinanceMigration.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "3.Schedule-Migration/New-FieldMappingFile.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "3.Schedule-Migration/Publish-MPPsToAltus.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "3.Schedule-Migration/Set-CustomFieldMapping.ps1"
            Arguments    = @("-customFieldName", "Diagnostic", "-localFieldTarget", "Text1")
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/-Run-Migration.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/Invoke-SharePointMigration.ps1"
            Arguments    = @("-SiteCollections", "diagnostic")
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/1-export-lists.ps1"
            Arguments    = @("-SiteCollectionUrl", "https://example.invalid", "-OutputFolder", $ProbePath, "-POLExportPath", $ProbePath)
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/2-import-lists.ps1"
            Arguments    = @("-D365Url", "https://example.invalid", "-DataFolder", $ProbePath, "-POLExportPath", $ProbePath)
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/3-export-documents.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "4.SharePoint-Migration/_Build-Package.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "5.TaskField-Migration/AltusTaskFieldMigration.ps1"
            Arguments    = @()
        }
        [PSCustomObject]@{
            RelativePath = "6.CalendarExceptions-Migration/1-export-calendar-exceptions.ps1"
            Arguments    = @("-InputFolder", $ProbePath, "-OutputFolder", $ProbePath)
        }
        [PSCustomObject]@{
            RelativePath = "6.CalendarExceptions-Migration/2-import-calendar-exceptions.ps1"
            Arguments    = @("-D365Url", "https://example.invalid", "-DataFolder", $ProbePath)
        }
    )
}

function Invoke-GuardedEntryPoint {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$EntryPoint
    )

    $scriptPath = Join-Path $script:RepositoryRoot $EntryPoint.RelativePath
    $output = (& $script:PowerShellPath -NoProfile -File $scriptPath @($EntryPoint.Arguments) 2>&1 | Out-String)

    [PSCustomObject]@{
        ExitCode = $LASTEXITCODE
        Output   = $output
    }
}

Describe "Debug-only execution boundary" {
    BeforeAll {
        $script:ProbePath = Join-Path ([System.IO.Path]::GetTempPath()) ("pol-debug-guard-" + [Guid]::NewGuid().ToString("N"))
        $script:EntryPoints = @(Get-GuardedEntryPoints -ProbePath $script:ProbePath)
    }

    It "defines the complete guarded entry-point allowlist" {
        $script:EntryPoints.Count | Should Be 20

        foreach ($entryPoint in $script:EntryPoints) {
            $scriptPath = Join-Path $script:RepositoryRoot $entryPoint.RelativePath
            Test-Path -LiteralPath $scriptPath -PathType Leaf | Should Be $true
        }
    }

    It "loads the shared guard before risky command nodes" {
        $riskyCommandPattern = '^(Connect-|Install-Module$|Read-Host$|Start-Process$|Invoke-RestMethod$|Invoke-WebRequest$|New-Object$|Set-Content$|Add-Content$|Out-File$|New-Item$|Export-|Import-)'

        foreach ($entryPoint in $script:EntryPoints) {
            $scriptPath = Join-Path $script:RepositoryRoot $entryPoint.RelativePath
            $content = Get-Content -LiteralPath $scriptPath -Raw
            $guardInvocation = 'Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path'
            $guardIndex = $content.IndexOf($guardInvocation, [System.StringComparison]::Ordinal)
            if ($guardIndex -lt 0) {
                throw "Guard invocation missing from $($entryPoint.RelativePath)."
            }

            $tokens = $null
            $parseErrors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
            if ($parseErrors.Count -gt 0) {
                throw "Parse error in $($entryPoint.RelativePath)."
            }

            $riskyCommands = $ast.FindAll({
                    param($node)
                    if ($node -isnot [System.Management.Automation.Language.CommandAst]) {
                        return $false
                    }

                    $commandName = $node.GetCommandName()
                    return $commandName -and $commandName -match $riskyCommandPattern
                }, $true)

            foreach ($riskyCommand in $riskyCommands) {
                if ($riskyCommand.Extent.StartOffset -lt $guardIndex) {
                    throw "Risky command '$($riskyCommand.GetCommandName())' precedes the guard in $($entryPoint.RelativePath)."
                }
            }
        }
    }

    It "refuses every entry point before creating output or authenticating" {
        foreach ($entryPoint in $script:EntryPoints) {
            $result = Invoke-GuardedEntryPoint -EntryPoint $entryPoint

            $result.ExitCode | Should Not Be 0
            $result.Output | Should Match "debugging reference only"
            $result.Output | Should Match "partners\.altus\.pro"
        }

        Test-Path -LiteralPath $script:ProbePath | Should Be $false
    }
}