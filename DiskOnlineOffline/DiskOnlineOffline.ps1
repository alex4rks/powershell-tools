# Disk Online Offline Changing Utility
# Kosarev Albert © 2016
# Select disk and change Online/Offline status 
# Selected disk cannot be system or removable
# Admin rights required
# Supported OS: Windows 7/8/10 with Powershell 4.0/5.0

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process "$psHome\powershell.exe" -Verb runAs -ArgumentList $arguments
break
}

Write-Host "Welcome to the Disk Online/Offline Status Changing Script!"
$DiskNum = (Get-Disk | Where {($_.IsSystem -eq $False) -and ($_.BusType -ne 'USB')}).Number
$diskstr = $disknum | & {$ofs=' and ';"$input"}
Write-Host "All available disks:" -NoNewLine
Get-Disk | Sort Number | ft Number, FriendlyName,  OperationalStatus, IsSystem, @{Name="Size (GB)";Expression={($_.Size / 1GB).ToString(".00")}}

Write-Host "You cannot change the state of the system disk / removable drive`nYou can change the state only of the disk(s): " $diskstr -BackgroundColor Black -ForegroundColor Yellow

Do {
	Write-Host "Write the disk number (q - quit): " -NoNewLine
	$DiskToChange = Read-Host
	if ($DiskTochange -like 'q*') { 
		Write-Host "Exiting.. Code: 3";
		Exit 3}
} while ($DiskNum -NotContains $DiskToChange)


$OfflineStatus = (get-disk -Number $DiskToChange).IsOffline;
$ChangedStatus = -not $OfflineStatus;
if ($OfflineStatus) {
	$StatusText = 'Offline'
	$NewStatusText = 'Online'}
else {
	$StatusText = 'Online'
	$NewStatusText = 'Offline'}

Write-Host "Currently disk " $DiskToChange "is "$StatusText".`nWould you like to change its state? [yes/no]: " -BackgroundColor Black -ForegroundColor Yellow -NoNewLine
$AcceptOptions = ('yes', 'no')
Do {
	$ChangeAccept = Read-Host
	if ($ChangeAccept -like 'n*') {
		Write-Host 'Exiting.. Code 3'
		Exit 3}
} While (-not ($ChangeAccept -like 'y*'))

Set-Disk -Number $DiskToChange -IsOffline $ChangedStatus;
Write-Host "New Status: "-NoNewLine
Write-Host $NewStatusText -BackgroundColor Black -ForegroundColor Yellow
Get-Disk -Number $disknum | ft Number, FriendlyName, OperationalStatus, IsSystem, @{Name="Size (GB)";Expression={($_.Size / 1GB).ToString(".00")}}
Write-host "Press any key to continue..."
[console]::ReadKey("NoEcho,IncludeKeyDown") | Out-Null
#Read-Host "Press any key.." 
