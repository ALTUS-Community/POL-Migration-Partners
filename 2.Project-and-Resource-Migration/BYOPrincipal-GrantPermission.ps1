

# Fail closed because this repository is a debugging reference only.
$debugOnlyGuardPath = Join-Path (Split-Path -Parent $PSScriptRoot) "DebugOnlyGuard.ps1"
if (-not (Test-Path -LiteralPath $debugOnlyGuardPath -PathType Leaf)) {
    throw "Execution blocked: DebugOnlyGuard.ps1 is unavailable."
}
. $debugOnlyGuardPath
Stop-PolMigrationExecution -ScriptPath $MyInvocation.MyCommand.Path

# Ensure required Graph modules are available
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites")
foreach ($moduleName in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "$moduleName not found. Installing for current user..." -ForegroundColor Yellow
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Error "Failed to install $moduleName`: $_"
            return
        }
    }

    try {
        Import-Module $moduleName -ErrorAction Stop
    } catch {
        Write-Error "Failed to import $moduleName`: $_"
        return
    }
}

# Connect with a user that that has Sites.FullControl.All as this is needed to make permissions changes
# (Global Admin, SharePoint Admin, or an app registration that has that permission)
$requiredScope = "Sites.FullControl.All"
$mgContext = Get-MgContext

if (-not $mgContext -or -not ($mgContext.Scopes -contains $requiredScope)) {
    # Write-Host "Connecting to Microsoft Graph (device code)..." -ForegroundColor Cyan
    # Write-Host "If a code is shown, open https://microsoft.com/devicelogin and enter it." -ForegroundColor DarkCyan
    # Connect-MgGraph -Scopes $requiredScope -UseDeviceCode -NoWelcome

    Connect-MgGraph -Scopes Sites.FullControl.All -NoWelcome

} else {
    Write-Host "Using existing Graph session for $($mgContext.Account)." -ForegroundColor Green
}

# -------------------------------------------------
# Configuration
# -------------------------------------------------
$defaultClientId = "YOUR-ENTRA-CLIENT-ID"
$displayName = "Enterprise App - Sites.Selected"   # optional but recommended
$defaultSiteUrl = "https://contoso.sharepoint.com/sites/pwa"
$role = "read"                              # read | write | manage | fullcontrol

$clientIdInput = Read-Host "Enter ClientID [$defaultClientId]"
$siteUrlInput = Read-Host "Enter SiteURL [$defaultSiteUrl]"

$clientId = if ([string]::IsNullOrWhiteSpace($clientIdInput)) { $defaultClientId } else { $clientIdInput.Trim() }
$siteUrl = if ([string]::IsNullOrWhiteSpace($siteUrlInput)) { $defaultSiteUrl } else { $siteUrlInput.Trim() }

# -------------------------------------------------
# Get the site
# -------------------------------------------------
# Convert the URL into the Graph site-id format
$siteIdLookup = $siteUrl -replace "https://", "" -replace "/sites/", ":/sites/"
$site = Get-MgSite -SiteId $siteIdLookup

if (-not $site) {
    Write-Error "Could not find site: $siteUrl"
    return
}

Write-Host "Found site: $($site.DisplayName) ($($site.WebUrl))" -ForegroundColor Green
Write-Host "Site ID   : $($site.Id)" -ForegroundColor Cyan

# -------------------------------------------------
# Grant the permission
# -------------------------------------------------
$params = @{
    roles               = @($role)
    grantedToIdentities = @(
        @{
            application = @{
                id          = $clientId
                displayName = $displayName
            }
        }
    )
}

try {
    $permission = New-MgSitePermission -SiteId $site.Id -BodyParameter $params
    Write-Host "`nSuccessfully granted '$role' permission." -ForegroundColor Green
    Write-Host "Permission ID: $($permission.Id)"
} catch {
    Write-Error "Failed to grant permission: $_"
}

# -------------------------------------------------
# Optional verification
# -------------------------------------------------
Write-Host "`nCurrent application permissions on the site:" -ForegroundColor Yellow
Get-MgSitePermission -SiteId $site.Id |
Where-Object { $_.GrantedToIdentities.Application } |
ForEach-Object {
    [PSCustomObject]@{
        PermissionId = $_.Id
        Roles        = $_.Roles -join ", "
        AppId        = $_.GrantedToIdentities.Application.Id
        AppName      = $_.GrantedToIdentities.Application.DisplayName
    }
} | Format-Table -AutoSize