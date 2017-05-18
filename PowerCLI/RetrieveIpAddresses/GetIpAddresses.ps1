# Retrieves IP addresses for all VApps within organisation on vSphere

# Author: Christopher Frantz
# Source: https://github.com/chrfrantz/PowerShell-scripts
# 
# Revision:
# 0.1 - Initial release

# General disclaimer regarding script execution:
# - Activate execution of scripts after understanding what they do (set-ExecutionPolicy RemoteSigned)
# - Deactivate execution of scripts if no longer required - leaving it activated is dangerous! (set-ExecutionPolicy Restricted)

# Instructions: 
# - Download scripts ConnectToVSphere.ps1, TemplateGlobals.ps1
# - Adjust values in TemplateGlobals.ps1 and username in ConnectToVSphere.ps1
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

#$vAppNetworkAdapters now contains names and IP addresses of all VApps in the organisation configured in TemplateGlobals.ps1

# Filter IP addresses based on VApp name and NICs
#$vAppNetworkAdapters = $vAppNetworkAdapters | where {$_.VAppName.startsWith($vAppIdentifier) -and $_.NIC -eq 'NIC0'}


