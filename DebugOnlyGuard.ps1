function Stop-PolMigrationExecution {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    $scriptName = Split-Path -Leaf $ScriptPath
    $message = @(
        "Execution blocked: this repository is a debugging reference only."
        "Do not run '$scriptName' from this repository."
        "Obtain the working scripts and migration tools from the approved Altus Partner Portal:"
        "https://partners.altus.pro/rooms/vdv9eco7129j5vxz5ta6z6p3?stage=wkxxnaeua9j1psh32tmafijh"
        "Run the working version there, then bring a redacted command, exact error, and question to an agent in this repository."
    ) -join [Environment]::NewLine

    throw $message
}