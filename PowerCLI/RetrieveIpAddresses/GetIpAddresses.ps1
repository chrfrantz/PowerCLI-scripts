# Retrieves IP addresses for all VApps within organisation on vSphere

# Author: Christopher Frantz
# Source: https://github.com/chrfrantz/PowerShell-scripts
# 
# Revision:
# 0.1 - Initial release

# Requirement: 
# - ConnectToVSphere.ps1, TemplateGlobals.ps1
# - Adjust values in TemplateGlobals.ps1
# - Connect using the script ConnectToVSphere.ps1 before running this script


$vAppNetworkAdapters = @()
foreach ($vApp in $vApps) {
        $vms = Get-CIVM -VApp $vApp
        foreach ($vm in $vms) {
                $networkAdapters = Get-CINetworkAdapter -VM $vm 
                foreach ($networkAdapter in $networkAdapters) {
                        $vAppNicInfo = New-Object "PSCustomObject"
                        $vAppNicInfo | Add-Member -MemberType NoteProperty -Name VAppName -Value $vApp.Name
                        $vAppNicInfo | Add-Member -MemberType NoteProperty -Name VMName   -Value $vm.Name
                        $vAppNicInfo | Add-Member -MemberType NoteProperty -Name NIC      -Value ("NIC" + $networkAdapter.Index)
                        $vAppNicInfo | Add-Member -MemberType NoteProperty -Name ExternalIP -Value $networkAdapter.IpAddress
                        $vAppNicInfo | Add-Member -MemberType NoteProperty -Name InternalIP -Value $networkAdapter.ExternalIpAddress

                        $vAppNetworkAdapters += $vAppNicInfo
                 }      
         }
}

# Filter IP addresses based on VApp name and NICs
#$vAppNetworkAdapters = $vAppNetworkAdapters | where {$_.VAppName.startsWith($vAppIdentifier) -and $_.NIC -eq 'NIC0'}


