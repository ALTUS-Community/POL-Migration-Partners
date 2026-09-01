<#
.SYNOPSIS
PS5-safe helper to relaunch scripts in PowerShell 7.

.DESCRIPTION
This file must remain compatible with Windows PowerShell 5.1 parser rules.
It is intentionally standalone and should be dot-sourced near the top of scripts.
#>

function Convert-SwitchLikeValuesToBoolean {
    <#
    .SYNOPSIS
    Normalizes deserialized switch-like objects to booleans.

    .DESCRIPTION
    JSON roundtrip can produce objects like @{ IsPresent = true } for switches.
    This function converts those values to plain [bool] for reliable parameter binding.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    foreach ($key in @($Parameters.Keys)) {
        $value = $Parameters[$key]
        if ($value -is [System.Collections.IDictionary] -and $value.Contains('IsPresent')) {
            $Parameters[$key] = [bool]$value['IsPresent']
        }
    }

    return $Parameters
}

function Get-PwshRelaunchCommand {
    <#
    .SYNOPSIS
    Builds the inline pwsh command used for relaunching with original bound parameters.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters
    )

    $boundParamsJson = $BoundParameters | ConvertTo-Json -Depth 10 -Compress
    $boundParamsB64 = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($boundParamsJson))
    $scriptPathEscaped = $ScriptPath.Replace("'", "''")

    $pwshCommand = @"
`$json = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$boundParamsB64'))
`$params = ConvertFrom-Json -InputObject `$json -AsHashtable
foreach (`$key in @(`$params.Keys)) {
    `$value = `$params[`$key]
    if (`$value -is [System.Collections.IDictionary] -and `$value.Contains('IsPresent')) {
        `$params[`$key] = [bool]`$value['IsPresent']
    }
}
& '$scriptPathEscaped' @params
exit `$LASTEXITCODE
"@

    return $pwshCommand
}

function Invoke-RelaunchInPwshIfNeeded {
    <#
    .SYNOPSIS
    Relaunches the current script in pwsh when running under Windows PowerShell 5.x.

    .OUTPUTS
    [bool] True when relaunch was attempted, otherwise false.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters,

        [Parameter(Mandatory = $false)]
        [switch]$NoExit
    )

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        return $false
    }

    $pwshPath = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Path
    if (-not $pwshPath) {
        Write-Host "ERROR: PowerShell 7 is required. Please install pwsh and retry." -ForegroundColor Red
        if (-not $NoExit) {
            exit 1
        }
        return $true
    }

    Write-Host "Relaunching script in PowerShell 7 (pwsh)..." -ForegroundColor Yellow
    $pwshCommand = Get-PwshRelaunchCommand -ScriptPath $ScriptPath -BoundParameters $BoundParameters

    & $pwshPath -NoProfile -ExecutionPolicy Bypass -Command $pwshCommand
    $childExitCode = $LASTEXITCODE

    if (-not $NoExit) {
        exit $childExitCode
    }

    return $true
}
