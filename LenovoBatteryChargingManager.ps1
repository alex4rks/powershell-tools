# Lenovo Battery Charging Manager
# © 2016
# This script can change start and stop charging tresholds
# The System must have 2 packages installed: 
# a) Lenovo Power Management Driver
# b) ThinkPad Settings Dependency
# Admin rights required to run the script
# Supported OS Windows 7/8/10 x86 and x64

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
	Write-Error "No Admin rights presented. Exiting.."
	Exit 5
}

Function SetPercentageValue { 
	param ([string]$ChargingParam, [int]$NewValue)
	Set-ItemProperty -Path $RegPath[0] -Name $ChargingParam -Type DWord -Value $NewValue
	Set-ItemProperty -Path $RegPath[1] -Name $ChargingParam -Type DWord -Value $NewValue
	Set-ItemProperty -Path $RegPath[2] -Name $ChargingParam -Type DWord -Value $NewValue
}

Function CheckChargingValue {
	param ($Value, $Path, $ValueName, $ValueRange1, $ValueRange2)
	$ValueRange = ($ValueRange1..$ValueRange2)
	if ($ValueRange -NotContains $Value) {
		Write-Error $Path" --> "$ValueName" contains wrong value: "$Value
	}
}


Write-Host "`nWelcome to the Lenovo Battery Charging Manager!`n" -Foreground Cyan
#
#Check OS architecture
$Arch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
if ($Arch -eq '64-bit') {
	$LenovoRegPath = "WOW6432Node\"
}
else {
    $LenovoRegPath = ""
}

# Check prerequisites
$ReqPaths = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Power Management Driver', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{08515684-CE49-47EF-B509-326A2E91BC5C}_is1')
$ReqNames = @('Lenovo Power Management Driver', 'ThinkPad Settings Dependency')

For ($i = 0; $i -lt 2; $i++) {
	($CheckVersion=(Get-ItemProperty -Path $ReqPaths[$i] -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion) | Out-Null; 
	if (-not $CheckVersion) {
		Write-Host "  "$ReqNames[$i]"`tnot found`tERROR" -Foreground Red
		$Exitcode = 5;
	} 
	else {
		Write-Host "  "$ReqNames[$i]"`t" -NoNewLine 
		Write-Host $CheckVersion"`tis OK  " -Foreground Yellow
	}
}
if ($ExitCode -ge 3) {
	Write-Host "You should install these packages to control notebook battery"
	Write-Host "`nExiting.. "
	Exit $ExitCode
}

#get battery information and display to host
$char = "|"
$data = Get-WmiObject -ClassName Win32_battery

#determine how much power remains
if ($data.EstimatedChargeRemaining -ge 80) {
    $color = "Green"
}
elseif ($data.EstimatedChargeRemaining -ge 40) {
    $color = "Yellow"
}
elseif ($data.EstimatedChargeRemaining -ge 20) {
    $color = "Magenta"
}
else {
    $color = "Red"
}
 
Write-Host "`r"
Write-Host $data.PSComputername" Battery" -ForegroundColor cyan 
Write-Host "Runtime    : " -NoNewline

if ($data.EstimatedRunTime -gt 1000000) {
	$RuntimeBigValue = " (Don`'t worry, it`'s OK)"
}
Write-Host "$(New-TimeSpan -minutes $data.EstimatedRunTime)"$RuntimeBigValue -ForegroundColor $color 
Write-Host "Remaining %: " -NoNewline
Write-host "$($data.EstimatedChargeRemaining)" -ForegroundColor $color
Write-Host "$($char*$($data.EstimatedChargeRemaining/2)) " -ForegroundColor $color
Write-Host "`r"

# Searching battery serial no.
$BatterySerial = (Get-ChildItem ("HKLM:\SOFTWARE\" + $LenovoRegPath + "Lenovo\PWRMGRV\ConfKeys\Data")  | Where {$_.PSChildName -cmatch "^[0-9][A-Z0-9]*$"}).PSChildName
if (-not ([string]::IsNullOrEmpty($BatterySerial))) {
	$RegPath = @("","","")
	$RegPath[0] = "HKLM:\SOFTWARE\" + $LenovoRegPath + "Lenovo\PWRMGRV\ConfKeys\Data\" + $BatterySerial
	$RegPath[1] = "HKLM:\SOFTWARE\" + $LenovoRegPath + "Lenovo\PWRMGRV\ConfKeys\Data"
	$RegPath[2] = "HKLM:\SOFTWARE\" + $LenovoRegPath + "Lenovo\PWRMGRV\Data"
} else { 
	Write-Host "Battery info is not found. Exiting.."
	Exit 3
}

ForEach ($Reg in $RegPath) {
	if (-not (Test-Path $Reg)) {
	Write-Host "Registry info is missed. Please check Registry settings or reinstall the software. Exiting.."
	Exit 3
	}
}

#Define Registry settings names
$RegSettingsNames = @("ChargeStartControl",
	"ChargeStartPercentage",
	"ChargeStopControl",
	"ChargeStopPercentage")

$TextValueOnOff = @('OFF', 'ON')
$TextValueOnOffNew = @('ON', 'OFF')
$ColorValueOnOff = @('Red', 'Green')
$ValueOnOffNew = @(1,0)
	
Function GetBatteryInfo() {
	# $RegPath[0] is the main path for reading data
	$Global:StartControl = ((Get-ItemProperty -Path $RegPath[0] -Name $RegSettingsNames[0]).($RegSettingsNames[0]))
	CheckChargingValue $Global:StartControl $RegPath[0] $RegSettingsNames[0] 0 1
	
	$Global:StartPercentage = ((Get-ItemProperty -Path $RegPath[0] -Name $RegSettingsNames[1]).($RegSettingsNames[1]))
	CheckChargingValue $Global:StartPercentage $RegPath[0] $RegSettingsNames[1] 0 100
	
	$Global:StopControl = ((Get-ItemProperty -Path $RegPath[0] -Name $RegSettingsNames[2]).($RegSettingsNames[2]))
	CheckChargingValue $Global:StopControl $RegPath[0] $RegSettingsNames[2] 0 1
	
	$Global:StopPercentage = ((Get-ItemProperty -Path $RegPath[0] -Name $RegSettingsNames[3]).($RegSettingsNames[3]))
	CheckChargingValue $Global:StopPercentage $RegPath[0] $RegSettingsNames[3] 0 100
}
Function ShowBatteryInfo() {
	Write-Host "`tBattery Charging Info" -ForegroundColor Cyan
	Write-Host "`tCharging Start Control`t`t " -NoNewLine 
	Write-Host " "$TextValueOnOff[$StartControl]" " -BackgroundColor Black -ForegroundColor $ColorValueOnOff[$StartControl]
	Write-Host "`tStart Charging At`t`t  "$StartPercentage"%"

	Write-Host "`tCharging Stop Control`t`t " -NoNewLine 
	Write-Host " "$TextValueOnOff[$StopControl]" " -BackgroundColor Black -ForegroundColor $ColorValueOnOff[$StopControl]
	Write-Host "`tStop Charging At`t`t  "$StopPercentage"%"
}
GetBatteryInfo
ShowBatteryInfo

Write-Host "`n`tCHOOSE COMMAND" -ForegroundColor Cyan
Write-Host "`t0. Show Battery Charging Info" -Foreground Yellow
Write-Host "`t1. ON/OFF Start Charging Control"
Write-Host "`t2. Change Start Percentage"
Write-Host "`t3. ON/OFF Stop Charging Control"
Write-Host "`t4. Change Stop Percentage"
Write-Host "`t5. Save Battery Report " -NoNewLine; 
Write-Host "*BONUS" -ForegroundColor Yellow
Write-Host "`t6. Exit"

$AvailableChoices = @('0','1','2','3','4','5','6')
$AvailablePercentage = (10..100)
$AvailableAnswersYesNo = ('')
Do {
	Write-Host "Write command (0-6): " -NoNewLine
	$Choice = Read-Host
	GetBatteryInfo
	Switch ($Choice) {
		0 { 
			ShowBatteryInfo
		}
		1 {
			Do {
				Write-Host "Are you sure to set Start Charging Control to "$TextValueOnOffNew[$StartControl] "? [yes/no] " -NoNewLine
				$StartAnswer = Read-Host 
			} While ( -not (($StartAnswer -like 'y*') -or ($StartAnswer -like 'n*')))
			if ($StartAnswer -like 'y*') {
				SetPercentageValue $RegSettingsNames[0] $ValueOnOffNew[$StartControl]
				Write-Host "Start Charging Control Set to "$TextValueOnOffNew[$StartControl] 
			}
		}
		2 { 
			Do{
				Write-Host "Write New Start Percentage Value 10-100: " -NoNewLine
				$NewStartPercentageValue = Read-Host
			} While ($AvailablePercentage -NotContains $NewStartPercentageValue)
			SetPercentageValue $RegSettingsNames[1] $NewStartPercentageValue
			Write-Host "Charging now starts at: "$NewStartPercentageValue" %"		
		}
		3 {
			Do {
				Write-Host "Are you sure to set Stop Charging Control to "$TextValueOnOffNew[$StopControl] "? [yes/no] " -NoNewLine
				$StopAnswer = Read-Host 
			} While ( -not (($StopAnswer -like 'y*') -or ($StopAnswer -like 'n*')))
			if ($StopAnswer -like 'y*') {
			SetPercentageValue $RegSettingsNames[2] $ValueOnOffNew[$StopControl]
			Write-Host "Start Charging Control Set to "$TextValueOnOffNew[$StopControl] 
			}
		}
		4  { 
			Do{
				Write-Host "Write New Stop Percentage Value 10-100: " -NoNewLine
				$NewStopPercentageValue = Read-Host
			} While ($AvailablePercentage -NotContains $NewStopPercentageValue)
			SetPercentageValue $RegSettingsNames[3] $NewStopPercentageValue
			Write-Host "Charging now stops at: "$NewStopPercentageValue" %"		
		}
		5 {
			$BatteryReportFileName = "C:\ProgramData\BatteryReport_" + (Get-Date -Format yyyy-MM-dd) + ".html"
			powercfg /batteryreport /output $BatteryReportFileName
			#get localized name of users group by well-known SID  
			$objSID = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")
			$objUsers = ($objSID.Translate( [System.Security.Principal.NTAccount])).Value
			# set full control to Users on the file
			$acl = Get-Acl $BatteryReportFileName
			$rule = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUsers, "FullControl", "Allow")
			$acl.SetAccessRule($rule)
			Set-Acl $BatteryReportFileName -AclObject $acl
		
		}
		6 {	
			Write-Host "`n"
			Write-Warning "`tSettings will be applied at next reboot!"
			Write-Host "Exiting.."
			Exit 3
		}
		default { Write-Warning "Please write number 1..6"; break}
	}
		
	

} while ($Choice -NotContains $AvailableChoices)
