# Set to $true only while debugging with Fiddler
$debug = $false
# Set this value to the Fiddler proxy URL configured on your computer
$proxyUrl = 'http://127.0.0.1:8888'

function Get-DataverseTenantId {
   param (
      [Parameter(Mandatory)]
      [string]
      $uri
   )
   # -SkipHttpErrorCheck returns the 401 response object rather than throwing,
   # letting us read the WWW-Authenticate header directly. This avoids PS5/PS7
   # differences in how HttpResponseException surfaces response headers.
   try {
      $response = Invoke-WebRequest -Uri ($uri + 'api/data/v9.2/') -UseBasicParsing -SkipHttpErrorCheck -ErrorAction Stop
      if ($response.StatusCode -eq 401) {
         # In PS7, Headers['WWW-Authenticate'] is IEnumerable<string>, so join to ensure we get a plain string.
         $wwwAuth = ($response.Headers['WWW-Authenticate'] -join ' ')
         if ($wwwAuth -match 'authorization_uri="https://login\.microsoftonline\.(com|us)/([^/"]+)') {
            return $Matches[2]
         }
         if ($wwwAuth -match 'authorization_uri=https://login\.microsoftonline\.(com|us)/([^/]+)') {
            return $Matches[2]
         }
      }
   } catch {
      # Tenant ID discovery is best-effort; Connect will fall back to login without -tenant
   }
   return $null
}

<#
.SYNOPSIS
Connects to Dataverse Web API using Azure authentication.

.DESCRIPTION
The Connect function uses the Get-AzAccessToken cmdlet to obtain an access token for the specified resource URI. 
It then sets the global variables baseHeaders and baseURI to be used for subsequent requests to the resource.

.PARAMETER uri
The resource URI to connect to. This parameter is mandatory.

.EXAMPLE
Connect -uri 'https://yourorg.crm.dynamics.com/'
This example connects to Dataverse environment and sets the baseHeaders and baseURI variables.
#>

function Connect {
   param (
      [Parameter(Mandatory)] 
      [String] 
      $uri,

      [ValidateSet('AzureCloud', 'AzureUSGovernment', 'AzureChinaCloud')]
      [string] $environment = 'AzureCloud',

      # Force re-authentication even if a cached Azure session exists.
      # Use this when switching between Dataverse environments/tenants.
      [Parameter(Mandatory = $false)]
      [switch]
      $Force
   )

   # Global variables overwrites parameter.
   if ($null -ne $global:Environment) {
      $environment = $global:Environment
   }

   if (-not [string]::IsNullOrWhiteSpace($environment)) {
      az cloud set --name $environment
   }

   # Discover the tenant ID from the Dataverse endpoint's 401 challenge.
   $tenantId = Get-DataverseTenantId -uri $uri

   # Check if already signed in to the correct tenant via Azure CLI.
   $currentTenant = az account show --query 'tenantId' -o tsv 2>$null
   $needsLogin = $Force -or ($LASTEXITCODE -ne 0) -or ($tenantId -and $currentTenant -ne $tenantId)

   if ($needsLogin) {
      # --allow-no-subscriptions is the key flag: it skips the tenant/subscription
      # selection prompt entirely, which appears when the user has no Azure subscriptions
      # in the target tenant (common for Dataverse-only users).
      $loginArgs = @('login', '--allow-no-subscriptions', '--only-show-errors')
      if ($tenantId) { $loginArgs += '--tenant', $tenantId }
      az @loginArgs | Out-Null
      if ($LASTEXITCODE -ne 0) {
         throw "Azure CLI login failed."
      }
   }

   # Get an access token for the Dataverse resource via Azure CLI.
   $tokenArgs = @('account', 'get-access-token', '--resource', $uri.TrimEnd('/'), '--query', 'accessToken', '-o', 'tsv', '--only-show-errors')
   if ($tenantId) { $tokenArgs += '--tenant', $tenantId }
   $token = (az @tokenArgs)
   if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($token)) {
      throw "Failed to acquire access token for $uri."
   }

   # Define common set of headers
   $global:baseHeaders = @{
      'Authorization'    = 'Bearer ' + $token
      'Accept'           = 'application/json'
      'OData-MaxVersion' = '4.0'
      'OData-Version'    = '4.0'
   }

   # Set baseURI
   $global:baseURI = $uri + 'api/data/v9.2/'
}


<#
.SYNOPSIS
Invokes a set of commands against the Dataverse Web API.

.DESCRIPTION
The Invoke-DataverseCommands function uses the Invoke-Command cmdlet to run a script block of commands against the Dataverse Web API. 
It handles any errors that may occur from the Dataverse API or the script itself.

.PARAMETER commands
The script block of commands to run against the Dataverse resource. This parameter is mandatory.

.EXAMPLE
Invoke-DataverseCommands -commands {
   # Get first account from Dataverse
   $accounts = (Get-Records `
      -setName 'accounts' `
      -query '?$select=name&$top=1').value

   $oldName = $accounts[0].name
   $newName = 'New Name'

   # Update the first account name to 'New Name'
   Set-ColumnValue `
      -setName 'accounts' `
      -id $accounts[0].accountid `
      -property 'name' `
      -value $newName

   Write-Host "First account name changed from '$oldName' to '$newName'"
}
This example invokes a script block that gets the first account from Dataverse and updates the name of the first account.
#>

function Invoke-DataverseCommands {
   param (
      [Parameter(Mandatory)] 
      $commands
   )
   try {
      Invoke-Command $commands -NoNewScope
   } catch [Microsoft.PowerShell.Commands.HttpResponseException] {
      Write-Host "An error occurred calling Dataverse:" -ForegroundColor Red
      $statuscode = [int]$_.Exception.StatusCode;
      $statusText = $_.Exception.StatusCode
      Write-Host "StatusCode: $statuscode ($statusText)"
      # Replaces escaped characters in the JSON
      [Regex]::Replace($_.ErrorDetails.Message, "\\[Uu]([0-9A-Fa-f]{4})", 
         { [char]::ToString([Convert]::ToInt32($args[0].Groups[1].Value, 16)) } )

   } catch {
      Write-Host "An error occurred in the script:" -ForegroundColor Red
      $_
   }
}

<#
.SYNOPSIS
Invokes a REST method with resilience to handle 429 errors.

.DESCRIPTION
The Invoke-ResilientRestMethod function uses the Invoke-RestMethod cmdlet to send an HTTP request to a RESTful web service. 
It handles any 429 errors (Too Many Requests) by retrying the request using the Retry-After header value as the retry interval.
It also handles 401 (Unauthorized) errors by re-acquiring the Azure CLI token and retrying the request once. 
It also supports using a proxy if the $debug variable is set to true.

.PARAMETER request
A hashtable of parameters to pass to the Invoke-RestMethod cmdlet. This parameter is mandatory.

.PARAMETER returnHeader
A boolean value that indicates whether to return the response headers instead of the response body. The default value is false.

.EXAMPLE
See the functions in the TableOperations.ps1 file for examples of using this function.
#>

function Invoke-ResilientRestMethod {
   param (
      [Parameter(Mandatory)] 
      $request,
      [bool]
      $returnHeader
   )

   if ($debug) {
      $request.Add('Proxy', $proxyUrl)
   }
   try {
      if ($returnHeader) {
         Invoke-RestMethod @request -ResponseHeadersVariable rhv | Out-Null
         return $rhv
      }
      Invoke-RestMethod @request
   } catch [Microsoft.PowerShell.Commands.HttpResponseException] {
      $statuscode = $_.Exception.Response.StatusCode
      if ($statuscode -eq 'Unauthorized') {
         # Token has expired — re-acquire via Azure CLI and retry once.
         Write-Host "  [$( Get-Date -Format 'HH:mm:ss')] 401 Unauthorized — refreshing token and retrying..." -ForegroundColor Yellow
         $rootUri = $global:baseURI -replace 'api/data/v9\.2/?$', ''
         Connect -uri $rootUri -Force
         # Rebuild the Authorization header on the outgoing request with the new token.
         $request.Headers = $global:baseHeaders.Clone()
         if ($request.ContainsKey('Body')) {
            $request.Headers['Content-Type'] = 'application/json'
         }
         if ($returnHeader) {
            Invoke-RestMethod @request -ResponseHeadersVariable rhv | Out-Null
            return $rhv
         }
         return Invoke-RestMethod @request
      } elseif ($statuscode -eq 'TooManyRequests') {
         if (!$request.ContainsKey('MaximumRetryCount')) {
            $request.Add('MaximumRetryCount', 3)
            # Don't need - RetryIntervalSec
            # When the failure code is 429 and the response includes the Retry-After property in its headers, 
            # the cmdlet uses that value for the retry interval, even if RetryIntervalSec is specified
         }
         # Will attempt retry up to 3 times
         if ($returnHeader) {
            Invoke-RestMethod @request -ResponseHeadersVariable rhv | Out-Null
            return $rhv
         }
         Invoke-RestMethod @request
      } else {
         throw $_
      }
   } catch {
      throw $_
   }
}
