# This script activates promiscuous mode and forged transmits for given virtual networks.
# Usage instructions:
# - Install PowerCLI environment on local machine and adjust paths shown below if necessary.
# - Modify server ($server) to match target server.
# - Modify network name pattern ($networkNameToModify) to match name of virtual networks for which promiscuous mode and forged transmits should be activated.
# - Adjust username to administrative username.
#
# Author: Christopher Frantz
#

$server = 'fthvc03'
$networkNameToModify = '*Internal*'

#http://community.spiceworks.com/scripts/show/2655-import-powercli-for-vmware-into-a-regular-ps-shell
function Import-PowerCLI {
	Add-PSSnapin vmware*
	if (Get-Item 'C:\Program Files (x86)' -ErrorAction SilentlyContinue) {
		. "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"
	}
	else {
		. "C:\Program Files\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"
	}
}

Import-PowerCLI

$PSModuleAutoloadingPreference = 'none'

Set-PowerCLIConfiguration -ProxyPolicy NoProxy -Confirm:$false

$pass = read-host -prompt "Enter password" -AsSecureString 

$pw = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))

Connect-VIServer -Server $server -Protocol https -User 'username' -Password $pw -ErrorAction Inquire

# Overwrite passwords
$pass = "done"
$pw = "done"

#Documentation on PortGroups
#https://pubs.vmware.com/vsphere-60/index.jsp?topic=%2Fcom.vmware.powercli.cmdletref.doc%2FSet-VDSecurityPolicy.html

write-host "Connected to $server."

$ct = Get-VDPortgroup | where { $_.Name -like $networkNameToModify } | Get-VDSecurityPolicy | where { $_.AllowPromiscuous -eq $false -or $_.ForgedTransmits -eq $false } | measure

$count = $ct.Count

Get-VDPortgroup | where { $_.Name -like $networkNameToModify } | Get-VDSecurityPolicy | where { $_.AllowPromiscuous -eq $false -or $_.ForgedTransmits -eq $false } | ft -AutoSize

$conf = read-host "We will set promiscuous mode and forged transmits for $count portgroups. Are you sure you want to continue? ('Yes', 'Y' or 'y')"
    
if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
    write-error "Execution has been aborted."
    Exit
}

Get-VDPortgroup | where { $_.Name -like $networkNameToModify } | Get-VDSecurityPolicy | 
    where { $_.AllowPromiscuous -eq $false -or $_.ForgedTransmits -eq $false } | Set-VDSecurityPolicy -AllowPromiscuous $true -ForgedTransmits $true -ErrorAction Stop

Get-VDPortgroup | where { $_.Name -like $networkNameToModify } | Get-VDSecurityPolicy