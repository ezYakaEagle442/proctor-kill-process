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
# Pre-requis
#
# PowerShell 7 utilise pwsh.exe
# powershell.exe lance toujours la version 5.1.
# 
#############################################################################

$PSH_VER="7.5.0"
Write-Host "PowerShell Version : $PSH_VER"
Write-Host ""
Write-Host "You must download & install https://github.com/PowerShell/PowerShell/releases/download/v$PSH_VER/PowerShell-$PSH_VER-win-x64.msi"
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
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running Tams KillO365Process (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where {($_.name -match "ms-teams") -or ($_.name -match "OfficeClickToRun") -or ($_.name -match "OneDrive") -or ($_.name -match "Outlook") -or ($_.name -match "winword") -or ($_.name -match "excel") -or ($_.name -match "POWERPNT") -or  ($_.name -match "slack") -or ($_.name -match "zoom") }
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
		Write-Log-Sub-Step "Searching for running Tams KillShellProcess (attempt #$attempt)..."
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
	Write-Log-Step "KillO365Process START"
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running Tams KillMiscProcess (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where {($_.name -match "adobe") -or ($_.name -match "NVIDIA") -or ($_.name -match "spoolsv") -or ($_.name -match "") -or ($_.name -match "ArmouryCrate") -or ($_.name -match "BrStMonW") -or  ($_.name -match "BrYNSvc") -or ($_.name -match "FileOpenManager64") -or ($_.name -match "FileOpenPIBroker") -or ($_.name -match "ssh-agent.exe") }
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
			Write-Log "All xxxx processes are now closed"
			Start-Sleep -Seconds 2
		}
		else {
			Write-Log "Found no running O365 processes"
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
	for ($attempt = 1; $attempt -le 10; $attempt++) {
		Write-Log-Sub-Step "Searching for running Tams KillVirtualizationProcess (attempt #$attempt)..."
        $RunningProcesses = Get-Process | Where {($_.name -match "vmms") -or ($_.name -match "vmcompute") -or ($_.name -match "vmware-authd") -or ($_.name -match "vmware-autostart") -or ($_.name -match "vmnetdhcp") -or ($_.name -match "vmnat") -or ($_.name -match "vmware-usbarbitrator64") -or  ($_.name -match "wsl")  -or  ($_.name -match "wslservice") }
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
	KillO365Process
	KillJavaProcess
	KillVirtualizationProcess
	KillMiscProcess
	KillBrowserProcess
	KillShellProcess
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