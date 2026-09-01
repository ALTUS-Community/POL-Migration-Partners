<#
.SYNOPSIS
Retries a script block up to 30 times if it fails with RPC_E_CALL_REJECTED error.

.DESCRIPTION
Executes the provided script block and automatically retries if an RPC_E_CALL_REJECTED error occurs.
Waits 1 second between retry attempts. Maximum 30 retries.
Original from ExportDraftAndPublishedAsXml.ps1 from https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online#scripts

.PARAMETER ScriptBlock
The script block to execute with retry logic.

.EXAMPLE
RunScriptBlockWithRetry -ScriptBlock { $WinProjApp.DisplayAlerts = $false }
#>
function RunScriptBlockWithRetry ($ScriptBlock) {
    $retryCount = 0
    $tryAgain = $true
    while ($tryAgain -and $retryCount -le 30) {
        try {
            Invoke-Command -ScriptBlock $ScriptBlock
            $tryAgain = $false
        } catch {
            if ($_.ToString().Contains("RPC_E_CALL_REJECTED")) {
                # retry on RPC_E_CALL_REJECTED error - with a second sleep
                Start-Sleep -Seconds 1
                $retryCount++
            } else {
                throw
            }
        }
    }
}

<#
.SYNOPSIS
Initializes a COM connection to an existing Microsoft Project instance.

.DESCRIPTION
Verifies the MS Project interop assembly is available and connects to an already-running
MS Project instance via COM interop. Sets $global:WinProjApp variable.
Original from ExportDraftAndPublishedAsXml.ps1 from https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online#scripts
Modified from Connect-WinProjToProjectServer as we no longer need to connect to Project Server.

.EXAMPLE
Open-WinProj

.NOTES
Microsoft Project must already be running before calling this function.
#>
function Open-WinProj {
    Write-Output "1. Verify existence of interop dll"
    $gac = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.MSProject').GlobalAssemblyCache
    if (-not $gac) {
        throw "Interop DLL not in GAC"
    }

    $projectProcess = [System.Diagnostics.Process]::GetProcessesByName("winproj");
    if (-not $projectProcess) {
        throw "Start MS Project before running this script."
    }

    Write-Output "2. Initializing the interop dll"
    $WinProjApp = New-Object -ComObject msproject.application
    RunScriptBlockWithRetry -ScriptBlock {
        $WinProjApp.DisplayAlerts = $false
    }

    $global:WinProjApp = $WinProjApp
}

<#
.SYNOPSIS
Closes the COM connection to Microsoft Project without saving.

.DESCRIPTION
Quits the Microsoft Project application via COM interop and releases the $global:WinProjApp variable.
Does not save any open projects.
Original from ExportDraftAndPublishedAsXml.ps1 from https://learn.microsoft.com/en-au/projectonline/export-user-data-from-project-online#scripts

.EXAMPLE
Close-WinProj
#>
function Close-WinProj() {
    if ($null -ne $global:WinProjApp) {
        RunScriptBlockWithRetry -ScriptBlock {
            $global:WinProjApp.Quit([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
        }

        $global:WinProjApp = $null
    }
}

<#
.SYNOPSIS
Retrieves the Altus for Project COM automation object.

.DESCRIPTION
Searches for the Altus for Project add-in in the MS Project COM add-ins collection
and returns its automation interface. Supports both ClickOnce and MSI installations.

.OUTPUTS
Returns the Altus for Project automation object.

.EXAMPLE
$automation = Get-AltusAutomationObject

.NOTES
Requires Open-WinProj to be called first to establish the COM connection.
#>
function Get-AltusAutomationObject() {
    if ($null -eq $global:WinProjApp) {
        throw "WinProj is not opened. Call Open-WinProj first."
    }

    Write-Host "Getting Altus Automation Object..."

    if (-not (Test-AltusAddinEnabled)) {
        throw "Altus for Project add-in is not enabled. Exiting."
    }

    $addin = RunScriptBlockWithRetry -ScriptBlock {
        # First search for ClickOnce AddIn 
        foreach ($a in $global:WinProjApp.COMAddIns) {
            if ($a.ProgId -eq "Altus" -or $a.ProgId -eq "Altus.Project" -or $a.ProgId -eq "Altus.ProjectAddIn") {
                return $a;
            }
        }

        # Second search for MSI installed AddIn 
        foreach ($a in $global:WinProjApp.COMAddIns) {
            if ($a.ProgId -eq "Sensei.AltusAddin") {
                return $a;
            }
        }

        return $null;
    }

    if ($null -eq $addin) {
        throw "Failed to find Altus for Project Automation."
    }

    if (-not $addin.Connect) {
        throw "Altus for Project Application COM Add-in is disabled."
    }

    $automation = $addin.Object;

    if ($null -eq $automation) {
        throw "Failed to find Altus for Project Automation object."
    }

    return $automation;
}

<#
.SYNOPSIS
Tests if the Altus for Project add-in is enabled and re-enables it if disabled.

.DESCRIPTION
Checks if the Altus for Project COM add-in is loaded and enabled in Microsoft Project.
If the add-in is disabled, attempts to re-enable it automatically.

.OUTPUTS
Returns $true if the add-in is enabled, $false otherwise.

.EXAMPLE
if (Test-AltusAddinEnabled) { Write-Host "Add-in is ready" }
#>
function Test-AltusAddinEnabled() {
    if ($null -eq $global:WinProjApp) {
        Write-Warning "WinProj is not opened. Cannot test add-in status."
        return $false
    }

    $addin = RunScriptBlockWithRetry -ScriptBlock {
        # First search for ClickOnce AddIn 
        foreach ($a in $global:WinProjApp.COMAddIns) {
            if ($a.ProgId -eq "Altus" -or $a.ProgId -eq "Altus.Project" -or $a.ProgId -eq "Altus.ProjectAddIn") {
                return $a
            }
        }

        # Second search for MSI installed AddIn 
        foreach ($a in $global:WinProjApp.COMAddIns) {
            if ($a.ProgId -eq "Sensei.AltusAddin") {
                return $a
            }
        }

        return $null
    }

    if ($null -eq $addin) {
        Write-Warning "Failed to find Altus for Project add-in."
        return $false
    }

    if (-not $addin.Connect) {
        try {
            RunScriptBlockWithRetry -ScriptBlock {
                $addin.Connect = $true
            }
            return $true
        } catch {
            Write-Warning "Altus for Project add-in is required for this script to run. Please ensure it is installed and enabled in MS Project."
            return $false
        }
    }

    return $true
}

<#
.SYNOPSIS
Closes all open projects in Microsoft Project.

.DESCRIPTION
Safely closes all projects currently open in MS Project using FileCloseAll.
Includes fallback logic to close projects individually if batch close fails.

.PARAMETER Save
If specified, saves projects before closing. Otherwise closes without saving.

.EXAMPLE
Close-AllOpenProjects

.EXAMPLE
Close-AllOpenProjects -Save
#>
function Close-AllOpenProjects {
    param(
        [switch]$Save
    )
    
    if ($null -eq $global:WinProjApp) {
        return
    }
    
    $saveType = if ($Save) { 'pjSave' } else { 'pjDoNotSave' }
    
    try {
        # Check if there are any open projects
        $projectCount = 0
        try {
            $projectCount = $global:WinProjApp.Projects.Count
        } catch {
            Write-Verbose "Could not get project count: $_"
            return
        }
        
        if ($projectCount -gt 0) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Closing $projectCount open project(s)..." -ForegroundColor Gray
            
            # Close all open projects
            [void](RunScriptBlockWithRetry -ScriptBlock {
                $global:WinProjApp.FileCloseAll([Microsoft.Office.Interop.MSProject.PjSaveType]::$saveType)
            })
        }
    } catch {
        Write-Verbose "Error closing projects (may already be closed): $_"
        
        # Fallback: try to close each project individually
        try {
            for ($i = $global:WinProjApp.Projects.Count; $i -ge 1; $i--) {
                try {
                    $global:WinProjApp.Projects.Item($i).Close($saveType)
                } catch {
                    Write-Verbose "Could not close project $($global:WinProjApp.Projects.Item($i).Name)"
                }
            }
        } catch {
            Write-Verbose "Fallback close failed"
        }
    }
}

<#
.SYNOPSIS
Tests if an error message indicates an RPC/COM connection failure.

.DESCRIPTION
Analyzes error messages to detect RPC and COM connection errors that typically
require a COM connection restart to recover.

.PARAMETER ErrorMessage
The error message string to analyze.

.OUTPUTS
Returns $true if the error is an RPC/COM connection error, $false otherwise.

.EXAMPLE
if (Test-IsRpcError -ErrorMessage $_.ToString()) { Restart-ComConnection }
#>
function Test-IsRpcError {
    param([string]$ErrorMessage)
    
    $rpcPatterns = @(
        '0x800706BA',           # RPC server is unavailable
        '0x800706BE',           # RPC call failed
        '0x800706BF',           # RPC call failed and did not execute
        '0x80010001',           # Call was rejected by callee (RPC_E_CALL_REJECTED)
        '0x80010108',           # Object invoked has disconnected from clients
        '0x800401FD',           # Object is not connected to server
        'RPC server is unavailable',
        'RPC call failed',
        'disconnected from its clients',
        'The object invoked has disconnected'
    )
    
    foreach ($pattern in $rpcPatterns) {
        if ($ErrorMessage -match [regex]::Escape($pattern)) {
            return $true
        }
    }
    return $false
}

<#
.SYNOPSIS
Restarts the COM connection to Microsoft Project without closing the application.

.DESCRIPTION
Releases the COM reference, performs garbage collection, waits for stabilization,
then re-establishes the COM connection to MS Project. Includes exponential backoff
retry logic. Returns a fresh Altus automation object.

.PARAMETER DelaySeconds
Initial delay in seconds before attempting reconnection. Default is 5 seconds.

.PARAMETER MaxAttempts
Maximum number of reconnection attempts. Default is 5 attempts.

.OUTPUTS
Returns a fresh Altus for Project automation object.

.EXAMPLE
$automation = Restart-ComConnection -DelaySeconds 10 -MaxAttempts 3
#>
function Restart-ComConnection {
    param(
        [int]$DelaySeconds = 5,
        [int]$MaxAttempts = 5
    )
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Refreshing COM connection to MS Project..." -ForegroundColor Magenta
    
    # Try to close any open projects first
    try {
        Close-AllOpenProjects
    } catch {
        Write-Verbose "Close-AllOpenProjects failed (expected if COM already broken): $_"
    }
    
    # Release the current COM reference without quitting Project
    $global:WinProjApp = $null
    
    # Force garbage collection to release COM references
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    # Wait for Project to stabilize
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Waiting $DelaySeconds seconds for MS Project to stabilize..." -ForegroundColor Gray
    Start-Sleep -Seconds $DelaySeconds
    
    # Retry logic for COM reconnection
    $attempt = 0
    $connected = $false
    $lastError = $null
    
    while (-not $connected -and $attempt -lt $MaxAttempts) {
        $attempt++
        
        if ($attempt -gt 1) {
            # Exponential backoff: 5, 10, 20, 40 seconds
            $backoffDelay = [Math]::Min(5 * [Math]::Pow(2, $attempt - 1), 60)
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] COM reconnect attempt $attempt of $MaxAttempts in $backoffDelay seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds $backoffDelay
            
            # GC before retry
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
        
        try {
            Write-Verbose "Creating new COM connection to existing Project instance (attempt $attempt)..."
            $WinProjApp = New-Object -ComObject msproject.application
            RunScriptBlockWithRetry -ScriptBlock {
                $WinProjApp.DisplayAlerts = $false
            }
            $global:WinProjApp = $WinProjApp
            $connected = $true
        } catch {
            $lastError = $_
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] COM reconnect attempt $attempt failed: $_" -ForegroundColor Red
            $global:WinProjApp = $null
        }
    }
    
    if (-not $connected) {
        throw "Failed to reconnect to MS Project after $MaxAttempts attempts. Last error: $lastError"
    }
    
    # Get fresh automation object
    $newAutomation = Get-AltusAutomationObject
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] COM connection refreshed successfully." -ForegroundColor Green
    
    return $newAutomation
}

<#
.SYNOPSIS
Publishes a project to Altus using the Altus for Project automation interface.

.DESCRIPTION
Opens an MPP file, publishes it to Altus, and optionally publishes a baseline.
The project is closed and saved after publication.

.PARAMETER projectPath
The full path to the MPP file to publish.

.PARAMETER automation
The Altus for Project automation object obtained from Get-AltusAutomationObject.

.PARAMETER isDebug
(Optional) Enable debug mode for additional logging.

.PARAMETER publishBaseline
(Optional) Switch to enable baseline publishing after project publish succeeds.

.PARAMETER baselineNumber
(Optional) The baseline number to publish (0-10). Only used if publishBaseline is true.

.PARAMETER isReportable
(Optional) Whether the baseline should be reportable. Only used if publishBaseline is true.

.OUTPUTS
Returns a PSCustomObject with Item1 and Item2 properties:
  - Item1: The project publish result object (or $null if publish failed)
  - Item2: The baseline publish result object (or $null if baseline not published)

.EXAMPLE
$automation = Get-AltusAutomationObject
$result = Invoke-AltusPublishProject -projectPath "C:\path\to\project.mpp" -automation $automation
$projectResult = $result.Item1
$baselineResult = $result.Item2
if ($projectResult.Success) { Write-Host "Project published successfully" }
#>
function Invoke-AltusPublishProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$projectPath,
        [Parameter(Mandatory = $true)]
        [object]$automation,
        [bool]$isDebug = $false,
        [switch]$publishBaseline,
        [int]$baselineNumber = 0,
        [bool]$isReportable = $true
    )
    if ($null -eq $global:WinProjApp) {
        throw "WinProj is not opened. Call Open-WinProj first."
    }

    if ($null -eq $automation) {
        throw "Failed to find Altus for Project Automation object."
    }

    $result = $null
    $baselineResult = $null
    $projectOpened = $false

    try {
        Write-Host "Opening project..."
        $WinProjApp.FileOpenEx($projectPath, $false)
        $projectOpened = $true
        $openedProject = $WinProjApp.ActiveProject
        
        # You can set check and force the project currency before publish if needed, here we set it to AUD as an example.
        Set-ProjectCurrency
        
        # Verify automation object is still connected before attempting publish
        if ($null -eq $automation) {
            throw "Automation object is null. The Altus for Project add-in may have been disabled."
        }
        
        Write-Host "AfP Automation object ready calling PublishProject..."
        
        try {
            $result = $automation.PublishProject()
            
            if ($null -eq $result) {
                Write-Warning "PublishProject() returned null. This may indicate:"
                Write-Warning "  - The add-in is not responding or has been disabled"
                Write-Warning "  - The project is not in a publishable state"
                Write-Warning "  - There are validation errors preventing publish"
                Write-Warning "Check the Altus for Project add-in is loaded and enabled in MS Project."
            }
        } catch {
            Write-Host "Exception during PublishProject(): $_" -ForegroundColor Red
            throw
        }
        
        $baselineMethodExists = $automation.PSObject.Methods.Name -contains "PublishBaseline"
        
        try {
            if ($publishBaseline -and $baselineMethodExists) {
                if ($null -ne $result -and $result.Success) {
                    Write-Host "Publish Project succeeded, proceeding to publish baseline..."
                    Write-Host "Publishing Baseline $baselineNumber (Reportable: $isReportable)..."
                    $baselineResult = $automation.PublishBaseline($baselineNumber, $isReportable)
                } else {
                    Write-Warning "Publish Project did not succeed, skipping baseline publish."
                }
            } elseif ($publishBaseline -and -not $baselineMethodExists) {
                Write-Host "Publish Baseline method not found, skipping baseline publish."
            } else {
                Write-Host "Baseline publishing not requested, skipping."
            }
        } catch {
            Write-Host "Error occurred while attempting to publish Baseline $baselineNumber`: $_" -ForegroundColor Red
        }

        # Save as XML copy if publish succeeded.
        # ActiveProject.Xml is non-functional and FileSaveAs pjXML is unsupported via COM automation
        # for OLE-format .mpp files from Project Online. Instead, build MSPDI-compatible XML directly
        # from the COM task object model, which is fully accessible regardless of file format.
        if ($null -ne $result -and $result.Success) {
            $xmlFileName = [System.IO.Path]::GetFileNameWithoutExtension($projectPath) + "_mpp.xml"
            $xmlPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($projectPath), $xmlFileName)
            Write-Host "Saving XML copy to: $xmlPath"
            try {
                $xml = [System.Xml.XmlDocument]::new()
                [void]$xml.AppendChild($xml.CreateXmlDeclaration('1.0', 'UTF-8', 'yes'))
                $projectNode = $xml.CreateElement('Project', 'http://schemas.microsoft.com/project')
                [void]$xml.AppendChild($projectNode)
                $tasksNode = $xml.CreateElement('Tasks', 'http://schemas.microsoft.com/project')
                [void]$projectNode.AppendChild($tasksNode)

                $taskCount = 0
                foreach ($task in $openedProject.Tasks) {
                    if ($null -eq $task) { continue }
                    $taskNode = $xml.CreateElement('Task', 'http://schemas.microsoft.com/project')

                    $uidNode = $xml.CreateElement('UID', 'http://schemas.microsoft.com/project')
                    $uidNode.InnerText = $task.UniqueID
                    [void]$taskNode.AppendChild($uidNode)

                    $guidNode = $xml.CreateElement('GUID', 'http://schemas.microsoft.com/project')
                    $guidNode.InnerText = $task.GUID
                    [void]$taskNode.AppendChild($guidNode)

                    $nameNode = $xml.CreateElement('Name', 'http://schemas.microsoft.com/project')
                    $nameNode.InnerText = $task.Name
                    [void]$taskNode.AppendChild($nameNode)

                    # Include Text30 (Altus Task ID)
                    $text30Val = $task.Text30
                    if (-not [string]::IsNullOrEmpty($text30Val)) {
                        $text30Node = $xml.CreateElement('Text30', 'http://schemas.microsoft.com/project')
                        $text30Node.InnerText = $text30Val
                        [void]$taskNode.AppendChild($text30Node)
                    }

                    [void]$tasksNode.AppendChild($taskNode)
                    $taskCount++
                }

                $xml.Save($xmlPath)
                Write-Host "Saved XML copy: $xmlFileName ($taskCount tasks)" -ForegroundColor Green
            } catch {
                Write-Host "  Warning: Could not save XML copy for $([System.IO.Path]::GetFileName($projectPath)): $_" -ForegroundColor Yellow
            }
        }

        Write-Host "Closing and Saving project..."
        $WinProjApp.FileCloseEx('pjSave')
        $projectOpened = $false
    } finally {
        # Ensure project is closed even if an error occurred
        if ($projectOpened) {
            try {
                Write-Host "Ensuring project is closed after error..." -ForegroundColor Yellow
                $WinProjApp.FileCloseEx('pjDoNotSave')
            } catch {
                Write-Verbose "Failed to close project in finally block: $_"
                # Try the aggressive close-all approach
                try {
                    Close-AllOpenProjects
                } catch {
                    Write-Verbose "Close-AllOpenProjects also failed: $_"
                }
            }
        }
    }

    # Return both results as an object with Item1 and Item2 properties (like a tuple)
    return [PSCustomObject]@{
        Item1 = $result
        Item2 = $baselineResult
    }
}

<#
.SYNOPSIS
Sets the currency code for the active project in Microsoft Project.

.DESCRIPTION
Checks if the active project's currency code matches the specified code.
If different, updates the currency code, symbol, and decimal digits.

.PARAMETER currencyCode
The currency code to set (e.g., "AUD", "USD"). Default is "AUD".

.PARAMETER currencySymbol
The currency symbol to display. Default is "$".

.PARAMETER currencyDigits
The number of decimal digits for currency values. Default is 2.

.EXAMPLE
Set-ProjectCurrency

.EXAMPLE
Set-ProjectCurrency -currencyCode "USD" -currencySymbol "$" -currencyDigits 2
#>
function Set-ProjectCurrency {
    param (
        [string]$currencyCode = "AUD",
        [string]$currencySymbol = "$",
        [int]$currencyDigits = 2
    )
    if ($null -eq $global:WinProjApp) {
        throw "WinProj is not opened. Call Open-WinProj first."
    }
    if ($WinProjApp.ActiveProject.CurrencyCode -ne $currencyCode) {
        Write-Host "Mismatch Detected in Currency Code. Updating Project Currency Code to $currencyCode"
        $WinProjApp.ActiveProject.CurrencyCode = $currencyCode
        $WinProjApp.ActiveProject.CurrencySymbol = $currencySymbol
        $WinProjApp.ActiveProject.CurrencyDigits = $currencyDigits
    }
}

<#
.SYNOPSIS
Finds a matching JSON metadata file for an MPP file using stem-based heuristics.

.DESCRIPTION
Searches for JSON files in the same directory as the MPP file that match by filename stem.
Supports flexible matching: exact stem match, JSON stem contains MPP stem, or vice versa.
If multiple JSON files match, returns the most recently modified one.

.PARAMETER MppFile
The MPP FileInfo object to find a matching JSON file for.

.OUTPUTS
Returns a FileInfo object for the matching JSON file, or $null if no match found.

.EXAMPLE
$mppFile = Get-Item "C:\Projects\Project_123_published.mpp"
$jsonFile = Find-JsonForMpp -MppFile $mppFile
if ($jsonFile) { Write-Host "Found: $($jsonFile.FullName)" }
#>
function Find-JsonForMpp {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$MppFile
    )
    
    $mppDir = $MppFile.DirectoryName
    $mppStem = [System.IO.Path]::GetFileNameWithoutExtension($MppFile.Name)
    
    # Search for JSON files in the same directory with matching stem patterns
    $jsonCandidates = Get-ChildItem -Path $mppDir -Filter "*.json" -File -ErrorAction SilentlyContinue | Where-Object {
        $jsonStem = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        # Match if: exact stem match, or JSON stem contains MPP stem, or MPP stem contains JSON stem
        ($jsonStem -eq $mppStem) -or 
        ($jsonStem -like "*$mppStem*") -or 
        ($mppStem -like "*$jsonStem*")
    }
    
    if (-not $jsonCandidates) {
        return $null
    }
    
    # If multiple matches, return the most recently modified
    if ($jsonCandidates.Count -gt 1) {
        return $jsonCandidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    
    return $jsonCandidates
}

<#
.SYNOPSIS
Parses ProjectLastPublishedDate from a JSON metadata file.

.DESCRIPTION
Reads a JSON file and extracts the ProjectLastPublishedDate from the expected structure:
NewDataSet.Project.ProjectLastPublishedDate
Returns $null if the file cannot be read, the path doesn't exist, or the date is missing/invalid.

.PARAMETER JsonPath
The full path to the JSON file to parse.

.OUTPUTS
Returns a DateTime object if the date was successfully parsed, or $null otherwise.

.EXAMPLE
$date = Get-ProjectLastPublishedDate -JsonPath "C:\Projects\Project_123_reporting_Project.json"
if ($date) { Write-Host "Last published: $($date.ToString('yyyy-MM-dd'))" }
#>
function Get-ProjectLastPublishedDate {
    param(
        [Parameter(Mandatory)]
        [string]$JsonPath
    )
    
    try {
        $jsonContent = Get-Content -Path $JsonPath -Raw -ErrorAction Stop | ConvertFrom-Json
        
        # Navigate nested structure: NewDataSet.Project.ProjectLastPublishedDate
        $dateStr = $null
        if ($jsonContent.NewDataSet -and $jsonContent.NewDataSet.Project) {
            $dateStr = $jsonContent.NewDataSet.Project.ProjectLastPublishedDate
        }
        
        if ([string]::IsNullOrWhiteSpace($dateStr)) {
            return $null
        }
        
        # Try to parse the date string
        $parsedDate = [DateTime]::MinValue
        if ([DateTime]::TryParse($dateStr, [ref]$parsedDate)) {
            return $parsedDate
        }
        
        return $null
    } catch {
        return $null
    }
}
<#
.SYNOPSIS
Tests if internet connectivity is available.

.DESCRIPTION
Performs a connectivity test against public internet endpoints over TCP/443.
Returns $true only if at least one public endpoint is reachable, which helps
differentiate internet connectivity from local intranet-only connectivity.

.PARAMETER TargetHost
Primary public host to test. Default is "www.microsoft.com".

.PARAMETER TimeoutSeconds
Timeout in seconds for the connectivity test. Default is 3 seconds.

.OUTPUTS
Returns $true if internet connectivity appears available, $false otherwise.

.EXAMPLE
if (Test-NetworkConnectivity) { Write-Host "Network is available" }

.EXAMPLE
if (-not (Test-NetworkConnectivity -TargetHost "www.bing.com" -TimeoutSeconds 3)) { Write-Host "Internet unavailable" }
#>
function Test-NetworkConnectivity {
    param(
        [string]$TargetHost = "www.microsoft.com",
        [int]$TimeoutSeconds = 3
    )
    Write-Verbose "Testing internet connectivity with timeout of $TimeoutSeconds seconds..."

    $publicHosts = @($TargetHost, 'www.google.com', 'www.cloudflare.com', '1.1.1.1') | Select-Object -Unique

    foreach ($testHost in $publicHosts) {
        $client = $null
        try {
            $client = New-Object System.Net.Sockets.TcpClient
            $connectTask = $client.ConnectAsync($testHost, 443)
            if ($connectTask.Wait($TimeoutSeconds * 1000)) {
                if ($client.Connected) {
                    Write-Verbose "Internet connectivity confirmed via $testHost:443"
                    return $true
                }
            } else {
                Write-Verbose "Timed out connecting to $testHost:443"
            }
        } catch {
            Write-Verbose "Failed connecting to $testHost:443 - $($_.Exception.Message)"
        } finally {
            if ($null -ne $client) {
                $client.Close()
                $client.Dispose()
            }
        }
    }

    return $false
}

<#
.SYNOPSIS
Performs diagnostic checks on network and application connectivity.

.DESCRIPTION
Performs a series of connectivity checks and returns diagnostic information useful
for troubleshooting network issues. Generates a report of what's working and what's not.

.OUTPUTS
Returns a PSCustomObject with network diagnostic information.

.EXAMPLE
$diags = Get-NetworkDiagnostics
if (-not $diags.NetworkAvailable) { Write-Host "Network is down!" }
Write-Host "Can resolve DNS: $($diags.DnsResolution)"
Write-Host "Loopback available: $($diags.LoopbackAvailable)"
#>
function Get-NetworkDiagnostics {
    $diagnostics = [PSCustomObject]@{
        Timestamp         = Get-Date
        NetworkAvailable  = $false
        DnsResolution     = $false
        LoopbackAvailable = $false
        LocalNetworkFound = $false
        ErrorDetails      = @()
    }

    # Local network interfaces up (excluding loopback)
    try {
        $upIfaces = @([System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
            $_.OperationalStatus -eq 'Up' -and $_.NetworkInterfaceType -ne [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback
        })
        $diagnostics.LocalNetworkFound = ($upIfaces.Count -gt 0)
        if (-not $diagnostics.LocalNetworkFound) {
            $diagnostics.ErrorDetails += 'No non-loopback network interfaces are Up.'
        }
    } catch {
        $diagnostics.ErrorDetails += "Failed to enumerate network interfaces: $($_.Exception.Message)"
    }

    # DNS resolution
    try {
        [void][System.Net.Dns]::GetHostEntry('www.microsoft.com')
        $diagnostics.DnsResolution = $true
    } catch {
        $diagnostics.ErrorDetails += "DNS resolution failed: $($_.Exception.Message)"
    }

    # Loopback reachability (ICMP)
    try {
        $diagnostics.LoopbackAvailable = Test-Connection -ComputerName '127.0.0.1' -Count 1 -Quiet -ErrorAction Stop
        if (-not $diagnostics.LoopbackAvailable) {
            $diagnostics.ErrorDetails += 'Loopback ping failed.'
        }
    } catch {
        $diagnostics.ErrorDetails += "Loopback check failed: $($_.Exception.Message)"
    }

    # Public internet reachability
    try {
        $diagnostics.NetworkAvailable = Test-NetworkConnectivity -TimeoutSeconds 3
        if (-not $diagnostics.NetworkAvailable) {
            $diagnostics.ErrorDetails += 'Public TCP connectivity check failed.'
        }
    } catch {
        $diagnostics.ErrorDetails += "Network connectivity test threw: $($_.Exception.Message)"
    }

    return $diagnostics
}
