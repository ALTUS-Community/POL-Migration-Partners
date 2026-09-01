#Version: 16.0.7715.1200
#usage
#   . \Users\myself\Documents\ExportDraftAndPublishedAsXML.ps1
#   Connect-WinProjToProjectServer -pwaUrl https://contoso.sharepoint.com/sites/pwa -siteId YOUR-PWA-SITE-ID
#   Export-PublishedProjectsAsXML -projectName ExampleProject -projectGuid YOUR-PROJECT-ID -exportFolder C:\Migration\Data\
#   Export-DraftProjectsAsXML -projectName ExampleProject -projectGuid YOUR-PROJECT-ID -exportFolder C:\Migration\Data\

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

function Connect-WinProjToProjectServer {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$pwaUrl,

		[Parameter(Mandatory = $true)]
		[string]$siteId
	)

	echo "1. verify existence of interop dll"
	$gac = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.MSProject').GlobalAssemblyCache
	if (-not $gac) {
		throw "interop dll not in gac"
	}

	echo "2. get the working directory for WinProj"
	$p = New-Object -ComObject msproject.application
	$path = $p.Path
	$p.Quit([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
	# echo $path

	echo "3. kill existing winproj processes"
	[System.Diagnostics.Process]::GetProcessesByName("winproj") | % { $_.Kill() }

	echo "4. start winproj connected to given site"
	$process = New-Object -TypeName System.Diagnostics.Process
	$process.StartInfo.FileName = "winproj.exe"
	$process.StartInfo.WorkingDirectory = $path
	$process.StartInfo.Arguments = [String]::Format("/s {0} /g {1}", $pwaUrl, $siteId)
	$process.Start() | Out-Null
	[System.Threading.Thread]::Sleep(5000)

	echo "5. initialize the interop dll"
	$WinProjApp = New-Object -ComObject msproject.application
	RunScriptBlockWithRetry -ScriptBlock {
		$WinProjApp.DisplayAlerts = $false
	}

	echo "6. verify connected profile"
	$retPwaUrl = RunScriptBlockWithRetry -ScriptBlock {
		return $WinProjApp.Profiles.ActiveProfile.Server
	};

	if ($retPwaUrl -contains $pwaUrl -or $pwaUrl -contains $retPwaUrl) {
		echo "successfully connected to $retPWaUrl"
	} else {
		echo "Connected to $retPwaUrl "
	}
	
	$global:WinProjApp = $WinProjApp
}

function Close-WinProj() {
	if ($global:WinProjApp -ne $null) {
		RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.Quit([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
		}

		$global:WinProjApp = $null
	}
}

function Reset-WinProjConnection {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$pwaUrl,

		[Parameter(Mandatory = $true)]
		[string]$siteId
	)
	
	Write-Warning "Attempting to reset WinProj connection due to COM error..."
	
	try {
		# Try to close any open files first
		if ($global:WinProjApp -ne $null) {
			try {
				$global:WinProjApp.FileCloseEx([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
			} catch {
				# Ignore errors during cleanup
			}
		}
		
		# Force close WinProj
		Close-WinProj
		
		# Kill any remaining WinProj processes
		Write-Host "Killing existing WinProj processes..."
		[System.Diagnostics.Process]::GetProcessesByName("winproj") | ForEach-Object { 
			try { $_.Kill(); $_.WaitForExit(5000) } catch { }
		}
		
		# Release COM object
		if ($global:WinProjApp -ne $null) {
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:WinProjApp) | Out-Null
			$global:WinProjApp = $null
		}
		
		# Force garbage collection
		[System.GC]::Collect()
		[System.GC]::WaitForPendingFinalizers()
		
		# Wait a moment before reconnecting
		Start-Sleep -Seconds 3
		
		# Reconnect
		Write-Host "Re-establishing connection to Project Server..."
		Connect-WinProjToProjectServer -pwaUrl $pwaUrl -siteId $siteId
		
		Write-Host "WinProj connection reset successfully"
		return $true
	} catch {
		Write-Error "Failed to reset WinProj connection: $_"
		return $false
	}
}

function Export-DraftProjectsAsXML {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$projectName,

		[Parameter(Mandatory = $true)]
		[Guid]$projectGuid,

		[Parameter(Mandatory = $true)]
		[string]$exportFolder
	)

	echo "1. make sure initialize the interop dll - $($projectName)"
	if ($global:WinProjApp -eq $null) {
		throw "Please call Connect-WinProjToProjectServer first"
	}

	echo "2. open file from draft store - $($projectName)"
	try {
		$openResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileOpenEx("<>\" + $projectName, $true)
		}
		if (-not $openResult) {
			throw "Failed to open file - $($projectName) - from draft store"
		}
	} catch {
		Write-Warning "Error opening draft project '$projectName': $_"
		# Check if this is a COM error that might have corrupted the connection
		if ($_.Exception.Message -match "unexpected error|RPC_E_|0x800706BA|0x80010001") {
			Write-Warning "Detected potential COM corruption. Connection may need reset."
		}
		throw
	}

	try {
		echo "3. save file as xml - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_draft.xml", $exportFolder, $projectName), "pjMPP", $false, $false, $true, $true, "", "", "", "MSProject.xml")
		}
		if (-not $saveResult) {
			throw "Failed to save - $($projectName) as XML"
		}
	
		echo "4. save file as mpp - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_draft.mpp", $exportFolder, $projectName), "pjMPP")
		}
		if (-not $saveResult) {
			throw "Failed to save -$($projectName) as MPP"
		}	
	} finally {
		echo "5. close file - $($projectName)"
		$closeResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileCloseEx('pjDoNotSave')
		}

		if (-not $closeResult) {
			echo "Failed to close $($projectName)"
		}
	}
}
function Export-PublishedProjectsAsXML {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$projectName,

		[Parameter(Mandatory = $true)]
		[Guid]$projectGuid,

		[Parameter(Mandatory = $true)]
		[string]$exportFolder
	)

	echo "1. make sure initialize the interop dll - $($projectName)"
	if ($global:WinProjApp -eq $null) {
		throw "Please call Connect-WinProjToProjectServer first"
	}

	echo "2. open file from published store - $($projectName)"
	try {
		$openResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileOpenEx("<1>\" + $projectName, $true)
		}
		if (-not $openResult) {
			throw "Failed to open file - $($projectName) - from published store"
		}
	} catch {
		Write-Warning "Error opening published project '$projectName': $_"
		# Check if this is a COM error that might have corrupted the connection
		if ($_.Exception.Message -match "unexpected error|RPC_E_|0x800706BA|0x80010001") {
			Write-Warning "Detected potential COM corruption. Connection may need reset."
		}
		throw
	}

	try {
		echo "3. save file as xml - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_published.xml", $exportFolder, $projectName), "pjMPP", $false, $false, $true, $true, "", "", "", "MSProject.xml")
		}
		if (-not $saveResult) {
			throw "Failed to save - $($projectName) as XML"
		}

		echo "4. save file as mpp - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_published.mpp", $exportFolder, $projectName), "pjMPP")
		}
		if (-not $saveResult) {
			throw "Failed to save - $($projectName) as MPP"
		}
	} finally {
		echo "5. close file - $($projectName)"
		$closeResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileCloseEx('pjDoNotSave')
		}
		if (-not $closeResult) {
			echo "Failed to close $($projectName)"
		}
	}
}
