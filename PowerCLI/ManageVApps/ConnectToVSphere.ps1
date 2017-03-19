

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

#Identify path this script is located in
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
# ... and import globals script
. "$scriptPath\TemplateGlobals.ps1"

Set-PowerCLIConfiguration -ProxyPolicy NoProxy -Confirm:$false

$pass = read-host -prompt "Enter password" -AsSecureString 

$pw = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass))

Connect-CIServer -Server $server -Port 443 -User 'username' -Password $pw -Org $org -ErrorAction Inquire

# Overwrite passwords
$pass = "done"
$pw = "done"

#Documentation
#http://pubs.vmware.com/vsphere-60/index.jsp#com.vmware.powercli.ug.doc/GUID-BC5860B2-5A5B-4518-A3F9-FF359FF18238.html

$myOrg = Get-Org $org

$vApps = Get-CIVApp -Org $myOrg

$myOrgVdc = Get-OrgVdc -Name $vdc

write-host "Connected to $server."
