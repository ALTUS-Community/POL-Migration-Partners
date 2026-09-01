# This is copied from Common.ps1 - CallOData function
# Just a more generic form so we can call other SP APIs where needed
function CallSPAPI([UserDataSoapClient] $Proxy, [string] $BaseUrl, [string] $Request, [bool] $OnPrem) {    
    if (-not $BaseUrl.EndsWith("/")) {
        $BaseUrl = $BaseUrl + "/"
    }

    # get baseuri (for authentication), odata base uri and actual request uri
    $baseUri = New-Object Uri $BaseUrl
    $oDataBaseUri = New-Object Uri -ArgumentList @($baseUri, "_api/")
    $oDataRequestUri = New-Object Uri -ArgumentList @($oDataBaseUri, $Request)

    $headers = @{
        Accept = "application/json"
    }

    $parameters = @{
        Uri             = $oDataRequestUri.AbsoluteUri
        Method          = "GET"
        Headers         = $headers
        UseBasicParsing = $true # for backcompat only
    }

    if (-not $OnPrem) {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

        # get auth cookie from the behavior attached to the proxy
        foreach ($behavior in $Proxy.ChannelFactory.Endpoint.Behaviors) {
            $userDataBehavior = $behavior -as [UserDataRequestBehavior]
            if ($userDataBehavior -ne $null) {
                $session.Cookies.SetCookies($baseUri, $behavior.AuthCookie)
                break
            }
        }

        $parameters.WebSession = $session
    } else {
        $parameters.UseDefaultCredentials = $true
    }

    return Invoke-RestMethod @parameters
}

# This is copied from ExportProjectUserContent.ps1 - so we can reuse it!
function ValidateDirectory([string] $OutputDirectory) {
    if (-not (Test-Path $OutputDirectory)) {
        Write-Host "Output directory does not exist, creating $OutputDirectory"
        mkdir $OutputDirectory > $null
        return $true
    } elseif (-not (Test-Path -PathType Container $OutputDirectory)) {
        Write-Error "Output location specified is not a directory. Please specify a directory."
        return $false
    } elseif ((Get-ChildItem $OutputDirectory | Measure-Object).Count -gt 0) {
        Write-Error "Output location is not empty"
        return $false
    }

    # Output location is an existing empty directory
    return $true
}

# Gets the ResourceUid of the current user
function GetCurrentUserResourceUid([UserDataSoapClient] $proxy, [string] $pwaUrl, [bool] $OnPrem) {
    $loginName = (CallSPAPI -Proxy $proxy -BaseUrl $pwaUrl -Request "Web/currentuser?`$select=LoginName" -OnPrem $OnPrem) | Select-Object -ExpandProperty LoginName

    $resourceResponse = CallOData -Proxy $proxy -BaseUrl $pwaUrl -Request `
    ("Resources?`$filter=ResourceNTAccount eq '{0}'&`$top=1" -f [System.Uri]::EscapeDataString($loginName)) `
        -OnPrem $OnPrem

    return $resourceResponse.value[0].ResourceId
}

# Repairs invalid JSON returned by the PSI proxy where double-quote characters inside
# string values are incorrectly escaped as """" (PowerShell/VB/CSV convention) instead
# of \" (JSON standard). The regex targets """" that are NOT followed by JSON structural
# terminators (comma, closing brace/bracket, or newline) which would indicate an empty
# string boundary rather than an in-value escaped quote.
function Repair-ProxyJson([string] $json) {
    return $json -replace '""(?=[^,\}\]\r\n])', '\"'
}