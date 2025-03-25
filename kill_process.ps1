Import-Module Microsoft.PowerShell.Utility

#############################################################################
#
# usage: pwsh.exe -NoProfile -ExecutionPolicy Bypass "./kill_process.ps1"
#
# PowerShell 7 utilise pwsh.exe
# powershell.exe lance toujours la version 5.1.
# 
#############################################################################

#############################################################################
#
# 
#
# PowerShell 7 utilise pwsh.exe
# powershell.exe lance toujours la version 5.1.
# 
#############################################################################

$PSH_VER="7.5.0"
Write-Host "PowerShell Version : $PSH_VER"
Write-Host ""
# Write-Host "You must download & install https://github.com/PowerShell/PowerShell/releases/download/v$PSH_VER/PowerShell-$PSH_VER-win-x64.msi"
Write-Host ""
$PSVersionTable.PSVersion
pwsh.exe -v
Write-Host ""

#Log File
$LogPath = "$env:windir\Temp"
$LogFile = "$LogPath\kill-process.log"

#Script
$ScriptName = $MyInvocation.MyCommand.Name
$ScriptPath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent

#Return codes
$ReturnCodes = @{"OK" = 0;
				"PIN-SYS-1" = 196; # This OS is not supported.
				"PIN_ERR_001_ACCESS_DENIED" = 1603; # Access to the path 'C:\' is denied.				
				}

#$OutputEncoding = New-Object -typename System.Text.UTF8Encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# $OutputEncoding = [Console]::OutputEncoding
#[Console]::OutputEncoding=[Text.Encoding]::Unicode

Function Write-Log {
	Param ([string]$logstring)
	Add-content $LogFile -value $logstring
	Write-Host $logstring
}

Function Write-Log-Step {
	Param ([string]$logstring)
	$Separator = "#" * ($logstring.length + 25)
	Write-Log $Separator
	Write-Log "$(Get-Date -Format G) - $logstring"
	Write-Log $Separator
}

Function Write-Log-Sub-Step {
	Param ([string]$logstring)
	$Separator = "-" * ($logstring.length + 25)
	Write-Log $Separator
	Write-Log "$(Get-Date -Format G) - $logstring"
	Write-Log $Separator
}

function CheckOS {
	Write-Log-Step "Check OS"
	$OS = Get-WmiObject -class Win32_OperatingSystem
	Write-Log "OS detected: $($OS.Caption) $($OS.OSArchitecture)"
	if (($OS.Version -match "10.0.26100") -and ($OS.OSArchitecture -match "64")){
		Write-Log "This OS is supported"
		# http://www.samlogic.net/articles/sysnative-folder-64-bit-windows.htm
		if ((Test-Path -Path $env:windir\SysNative) -eq $true) {
			Write-Log "32-bit environment of execution detected!"
			return 32 ;
		}
		else {
			Write-Log "This x64 OS is supported"
			return 64 ;
		}
	}
	else {
		if (($OS.Version -match "10.0.26100") -and ($OS.OSArchitecture -match "32")) {
			Write-Log "This x86 OS is supported"
			return 32 ;
		}
		else {
			Write-Log "This OS is not supported."
			TerminateScript "PIN-SYS-1"
		}
	}
}


#*********************************************************************
# Kill Browser Process
#*********************************************************************
function KillBrowserProcess() {
	Write-Log-Step "KillBrowserProcess START"
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running BROWSER processes (attempt #$attempt)..."
		# msedge
        $RunningProcesses = Get-Process | Where {($_.name -match "iexplore") -or ($_.name -match "firefox") -or ($_.name -match "chrome")}
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running BROWSER processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}
			Write-Log-Sub-Step "Closing all running BROWSER processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All Browser processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running BROWSER processes"
			Break
		}
	}
	Write-Log-Step "KillBrowserProcess END"
}

#*********************************************************************
# Kill Java Process
#*********************************************************************
function KillJavaProcess() {
	Write-Log-Step "KillJavaProcess START"
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running BROWSER processes (attempt #$attempt)..."
		$RunningProcesses = Get-Process | Where {($_.name -match "javaw") -or ($_.name -match "javaws") -or ($_.name -match "jp2launcher") -or ($_.name -match "jusched")}
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running Java processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}
			Write-Log-Sub-Step "Closing all running Java processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All Java processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running Java processes"
			Break
		}
	}
	Write-Log-Step "KillJavaProcess END"
}

#*********************************************************************
# Kill O365 Process
#*********************************************************************
function KillO365Process() {
	Write-Log-Step "KillO365Process START"

	$service = Get-Service -Name "OneDrive Updater Service" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping OfficeClickToRun service..."
		Stop-Service -Name "OneDrive Updater Service" -Force
		Write-Log "OneDrive Updater Service stopped"
	}

	$service = Get-Service -Name "Microsoft.SharePoint" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping Microsoft.SharePoint service..."
		Stop-Service -Name "Microsoft.SharePoint Service" -Force
		Write-Log "Microsoft.SharePoint Service stopped"
	}
	

	$service = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping OfficeClickToRun service..."
		Stop-Service -Name ClickToRunSvc -Force
		Write-Log "OfficeClickToRun service stopped"
	}

	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running KillO365Process (attempt #$attempt)..."
		# /!\ Outlook process is identified by "olk" process NOT "RuntimeBroker"
        $RunningProcesses = Get-Process | Where {($_.name -match "olk") -or ($_.name -match "Outlook") -or ($_.name -match "ms-teams") -or ($_.name -match "OfficeClickToRun") -or ($_.name -match "OneDrive") -or ($_.name -match "winword") -or ($_.name -match "excel") -or ($_.name -match "POWERPNT") }
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running O365 processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}

			Write-Log-Sub-Step "Closing all running O365 processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All 0365 processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running O365 processes"
			Break
		}
	}
	Write-Log-Step "KillO365Process END"
}

#*********************************************************************
# Kill Shell Process
#*********************************************************************
function KillShellProcess() {
	Write-Log-Step "KillShellProcess START"
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for CMD (attempt #$attempt)..."
        # $RunningProcesses = Get-Process | Where {($_.name -match "wsl") -or ($_.name -match "CMD") -or ($_.name -match "powershell") -or ($_.name -match "pwsh") }
        $RunningProcesses = Get-Process | Where { ($_.name -match "CMD") }
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running Shell processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}
			Write-Log-Sub-Step "Closing all running Shell processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All Shell processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running Shell processes"
			Break
		}
	}
	Write-Log-Step "KillShellProcess END"
}

#*********************************************************************
# Kill Miscellanous Process
#*********************************************************************
function KillMiscProcess() {
	Write-Log-Step "KillMiscProcess START"

	$service = Get-Service -Name Spooler -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping Spooler service..."
		Stop-Service -Name Spooler -Force
		Write-Log "Spooler service stopped"
	}

	$service = Get-Service -Name FileOpenManager -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping FileOpenManager service..."
		Stop-Service -Name FileOpenManager -Force
		Write-Log "FileOpenManager service stopped"
	}

	$service = Get-Service -Name NvContainerLocalSystem -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping NvContainerLocalSystem  service..."
		Stop-Service -Name NvContainerLocalSystem  -Force
		Write-Log "NvContainerLocalSystem service stopped"
	}

	
	$service = Get-Service -Name BcastDVRUserService_49a0a6 -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping BcastDVRUserService_49a0a6  service..."
		Stop-Service -Name NvContainerLocalSystem  -Force
		Write-Log "BcastDVRUserService_49a0a6 service stopped"
	}

	$service = Get-Service -Name ssh-agent -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping ssh-agent  service..."
		Stop-Service -Name ssh-agent  -Force
		Write-Log "ssh-agent service stopped"
	}

	$service = Get-Service -Name ArmouryCrateControlInterface -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping ArmouryCrateControlInterface  service..."
		Stop-Service -Name ArmouryCrateControlInterface -Force
		Write-Log "ArmouryCrateControlInterface service stopped"
	}

	$service = Get-Service -Name ArmouryCrateService -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping ArmouryCrateService  service..."
		Stop-Service -Name ArmouryCrateService -Force
		Write-Log "ArmouryCrateService service stopped"
	}

	$service = Get-Service -Name BrStMonW -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping BrStMonW  service..."
		Stop-Service -Name BrStMonW -Force
		Write-Log "BrStMonW service stopped"
	}	

	$service = Get-Service -Name BrYNSvc -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping BrYNSvc  service..."
		Stop-Service -Name BrYNSvc -Force
		Write-Log "BrYNSvc service stopped"
	}	

	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running Tams KillMiscProcess (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where { ($_.name -match "Code") -or  ($_.name -match "slack") -or ($_.name -match "zoom") -or ($_.name -match "adobe") -or ($_.name -match "acrobat") -or ($_.name -match "NVIDIA") -or ($_.name -match "NVIDIA Broadcast") -or ($_.name -match "nvcontainer")  -or ($_.name -match "NVDisplay.Container") -or ($_.name -match "spoolsv") -or ($_.name -match "ArmouryCrate") -or ($_.name -match "BrStMonW") -or  ($_.name -match "BrYNSvc") -or ($_.name -match "FileOpenManager64") -or ($_.name -match "FileOpenPIBroker") -or ($_.name -match "ssh-agent.exe") }
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running MISC processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}

			Write-Log-Sub-Step "Closing all running MISC processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All MISC processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running MISC processes"
			Break
		}
	}
	Write-Log-Step "KillMiscProcess END"
}

#*********************************************************************
# Kill VMWare & Hyper-V Process
#*********************************************************************
function KillVirtualizationProcess() {
	Write-Log-Step "KillVirtualizationProcess START"

	$service = Get-Service -Name "AppVClient" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "AppVClient" service ..."
		Stop-Service -Name "AppVClient" -Force
		Write-Log "AppVClient service stopped"
	}
	
	$service = Get-Service -Name "com.docker.service" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "com.docker.service" service ..."
		Stop-Service -Name "com.docker.service" -Force
		Write-Log "com.docker.service service stopped"
	}

	$service = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmcompute" service ..."
		Stop-Service -Name "vmcompute" -Force
		Write-Log "vmcompute service stopped"
	}

	$service = Get-Service -Name "vmicguestinterface" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicguestinterface" service ..."
		Stop-Service -Name "vmicguestinterface " -Force
		Write-Log "vmicguestinterface service stopped"
	}

	$service = Get-Service -Name "vmicheartbeat" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicheartbeat" service ..."
		Stop-Service -Name "vmicheartbeat" -Force
		Write-Log "vmicheartbeat service stopped"
	}

	$service = Get-Service -Name "vmicrdv" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicrdv" service ..."
		Stop-Service -Name "vmicrdv" -Force
		Write-Log "vmicrdv service stopped"
	}
	$service = Get-Service -Name "vmicshutdown" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicshutdown" service..."
		Stop-Service -Name "vmicshutdown" -Force
		Write-Log "vmicshutdown service stopped"
	}
	$service = Get-Service -Name "vmictimesync" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmictimesync" service..."
		Stop-Service -Name "vmictimesync" -Force
		Write-Log "vmictimesync service stopped"
	}

	$service = Get-Service -Name "vmicvmsession" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicvmsession" service..."
		Stop-Service -Name "vmicvmsession" -Force
		Write-Log "vmicvmsession service stopped"
	}

	$service = Get-Service -Name "vmicvss" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmicvss" service..."
		Stop-Service -Name "vmicvss" -Force
		Write-Log "vmicvss  service stopped"
	}

	$service = Get-Service -Name "vmms" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "vmms" service ..."
		Stop-Service -Name "vmms" -Force
		Write-Log "vmms  service stopped"
	}

	$service = Get-Service -Name "VMnetDHCP" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "VMnetDHCP" service..."
		Stop-Service -Name "VMnetDHCP" -Force
		Write-Log "VMnetDHCP service stopped"
	}

	$service = Get-Service -Name "VMUSBArbService" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "VMUSBArbService" service..."
		Stop-Service -Name "VMUSBArbService" -Force
		Write-Log "VMUSBArbService service stopped"
	}

	$service = Get-Service -Name "VMware NAT Service" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "VMware NAT Service" service..."
		Stop-Service -Name "VMware NAT Service" -Force
		Write-Log "VMware NAT Service service stopped"
	}

	$service = Get-Service -Name "VmwareAutostartService" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "VmwareAutostartService" service..."
		Stop-Service -Name "VmwareAutostartService" -Force
		Write-Log "VmwareAutostartService service stopped"
	}

	$service = Get-Service -Name "WSLService" -ErrorAction SilentlyContinue
	if ($service.Status -eq "Running") {
		Write-Log "Stopping "WSLService" service ..."
		Stop-Service -Name "WSLService" -Force
		Write-Log "WSLService service stopped"
	}

	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running Tams KillVirtualizationProcess (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where {($_.name -match "vmms") -or ($_.name -match "vmcompute") -or ($_.name -match "vmware-authd") -or ($_.name -match "vmware-autostart") -or ($_.name -match "vmnetdhcp") -or ($_.name -match "vmnat") -or ($_.name -match "vmware-usbarbitrator64") -or ($_.name -match "wsl") -or ($_.name -match "com.docker.build") -or ($_.name -match "wslservice") }
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running Virtualization processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}

			Write-Log-Sub-Step "Closing all running Virtualization processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All Virtualization processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running Virtualization processes"
			Break
		}
	}
	Write-Log-Step "KillVirtualizationProcess END"
}

function KillProcess() {
	Write-Log-Step "KillProcess START"

	Write-Log-Step "List of ALL services"
	Write-Log-Step ""
	# Get-Service | Where-Object { $_ -match "OneDrive" }
	# Get-Service | Format-Table -AutoSize
	# Get-Process | Format-Table -AutoSize
	# $(Get-Process).processname | Where-Object { $_ -match "olk" }
	# Get-Service | Out-GridView

	KillO365Process
	KillJavaProcess
	KillVirtualizationProcess
	KillMiscProcess
	KillShellProcess
	KillBrowserProcess
	Write-Log-Step "KillProcess END"
}

[int]$osBits = CheckOS
Write-Host ""
Write-Host "This script will kill all running BROWSER processes, assuming Edge is used."
Write-Host ""

Write-Host "Have you read carefully the README file ?[Yes/No]: "
$READ_CHECK = Read-Host
Write-Host ""

if ($READ_CHECK -eq 'y' -or $READ_CHECK -eq 'Yes') {
    Write-Log-Step MAIN START
    Write-Host ""
    
    Write-Log-Sub-Step "ScriptName: $ScriptName"
    Write-Log-Sub-Step "ScriptPath: $ScriptPath"

    KillProcess

	# finally kill itself
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running PowerShell Process (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where { ($_.name -match "powershell") -or ($_.name -match "pwsh") }
        if ($RunningProcesses.Count -gt 0) {
			Write-Log "Found the following running PowerShell processes:"
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log $xProcess.Name
			}
			Write-Log-Sub-Step "Closing all running PowerShell processes..."
			ForEach ($xProcess in $RunningProcesses) {
				Write-Log "$(Get-Date -Format G) - Stopping ""$($xProcess.Name)"" process..."
				$xProcess | Stop-Process -Force
				Write-Log "$(Get-Date -Format G) - Process stopped"
			}
			Write-Log "All PowerShell processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running PowerShell processes"
			Break
		}
	}

    Write-Host ""
    Write-Log-Step MAIN END
} else {
    Write-Log-Step "You should read carefully the README file ..."
}

exit $LastExitCode