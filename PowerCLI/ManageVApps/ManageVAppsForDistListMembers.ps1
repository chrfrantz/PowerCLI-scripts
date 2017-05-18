# This script creates, deletes, or modifies given vApps based on Outlook distribution lists.
#
# Author: Christopher Frantz
# Source: https://github.com/chrfrantz/PowerShell-scripts
#
# Revision:
# 0.1 - Initial release
#
# General disclaimer regarding script execution:
# - Activate execution of scripts after understanding what they do (Command: set-ExecutionPolicy RemoteSigned)
# - Deactivate execution of scripts if no longer required - leaving it activated is dangerous! (Command: set-ExecutionPolicy Restricted)
#
# Instructions:
# - Configure VCD and Org settings in TemplateGlobals.ps1
# - Adjust username in ConnectToVSphere.ps1
# - Run ConnectToVSphere.ps1

# Identify path this script is located in
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
# ... and import globals script (only uses general variables)
. "$scriptPath\TemplateGlobals.ps1"

# Indicates whether Outlook Distribution list should be read into $table variable (good if run repeatedly)
$readDistListToTable = 
$true
#$false

# Tests for user existence independent of VApp creation
$testUserExistence = 
#$true
$false

# Creates vApp for users in Distribution List
$createVAppsForUsers = 
$true
#$false

# Deletes VApps for users in Distribution List
$deleteVAppsForUsers = 
#$true
$false

# Indicates whether permissions should be set.
$setPermissions = 
$true
#$false

# Indicates whether created vApps should be started.
$startVApps = 
$true
#$false

# Indicates whether VMs should be restarted after vApp start (e.g. to activate customised settings - even if no explicit customisation occurs).
$restartVMsAfterVAppStart = 
$true
#$false

# Indicates whether VMs should be restarted after vApp start, even if existing and despite customisation! - Careful: work can be lost
$restartVMsEvenIfExisting = 
#$true
$false

# Indicates whether running vApps should be stopped.
$stopVApps = 
#$true
$false

# Indicates whether hostnames should be customized.
$customizeHostnames = 
#$true
$false

# Indicates whether usernames are appended (_username) to custom hostnames.
$appendUserToCustomHostname = 
$true
#$false

# Distribution list name in Outlook
$distListName = 
'Outlook distribution list name'

# Template to build vApp from
$vAppTemplate = 
'vAppTemplateName'

# Prefix for created vApps
$vAppPrefix = 
'vAppPrefix'

# Suffix for created vApps
$vAppSuffix = ''

# Permission set for users -- set to 'None' to remove any access #CAN BE USED IN CONJUNCTION WITH CREATION IF PERMISSIONS OF EXISTING VMs SHOULD BE CHANGED (will only change permissions)
$permission = 
#'Read' # does not allow starting of VMs (only good for SBAs)
'ReadWrite'# allows network configuration (general default for labs)
#'FullControl' # allows sharing with others (good for self-guided projects)
#'None' # removes any access (invisible to user)

# HashTable containing all hostnames to be changed - key is Virtual machine names (not hostnames!), value is target hostname
$VmNamesToChange = 
@{“currentName“ = "targetName"}

#Indicates whether system should prompt to extend attempts to restart in case number of retries is exhausted.
$promptForFailedRestarts = 
#$true
$false

# Array of VMs for which restart failed.
$failedVMRestarts = @()

#try to get VCloud users - will fail if not connected

# Try a command that depends on vCloud connectivity and redirect stdout to null, since output not relevant.
get-CIUser > $null
if(!$?){
    Write-Host "Not connected vCloud. Operation aborted."
    Exit
}

if($customizeHostnames -and !$startVApps){
    Write-Error "Note that the customization of hostnames requires machines to restart. Starting of VApps should therefore be activated."
    Exit
}

if($customizeHostnames) {

    if(!$deleteVAppsForUsers){
        $stringified = $VmNamesToChange | out-string

        $conf = read-host "You requested the customization of VM hostnames ($stringified) including restart. Please confirm these modifications ('Yes', 'Y' or 'y')."
    
        if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
            write-error "Execution has been aborted."
            Exit
        }
    } else {
        Write-host "Customization of names is ignored, since VMs are to be deleted."
        $customizeHostnames = $false
    }
}

if($customizeHostnames -and $restartVMsAfterVAppStart) {
    write-host "You activated customization. Note: If customization is activated, individual VMs will implicitly be restarted as part of that process (whether or not their (re)start is explicitly requested)."
}

if($restartVMsEvenIfExisting) {
    $conf = read-host "You requested the restart of all VMs in any running or non-running vApps, not just for customisation. Note: You can lose work in existing VMs! Please confirm this intent ('Yes', 'Y' or 'y')."
    
    if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
        write-error "Execution has been aborted."
        Exit
    }
}

if($setPermissions -and $deleteVAppsForUsers) {
    write-host "Setting of permissions is ignored, since VMs are to be deleted."
    $setPermissions = $false
}

if($createVAppsForUsers -and $deleteVAppsForUsers) {
    Write-Error "You have specified the deletion and creation of vApps at the same time. Decide for either one and rerun."
    Exit
}

if(!$deleteVAppsForUsers -and $stopVApps) {
    
    $conf = read-host "Will stop all running vApps at the end of creation/modification (even if starting is activated!). Are you sure you want to continue? ('Yes', 'Y' or 'y')"
    
    if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
        write-error "Execution has been aborted."
        Exit
    }
    if($createVAppsForUsers) {
        
        $conf = read-host "Note that creation of vApps is activated. The script will attempt to create vApps before stopping them. (You could switch off the creation if you just want to stop vApps.) Are you sure you want to continue? ('Yes', 'Y' or 'y')"
        
        if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
            write-error "Execution has been aborted."
            Exit
        }
    }
    $startVApps = $false
}

if($deleteVAppsForUsers) {
    
    $conf = read-host "You requested the deletion of vApps (starting with '$vAppPrefix'). Are you sure you want to continue? ('Yes', 'Y' or 'y')"

    if($conf -ne 'Y' -and $conf -ne 'Yes' -and $conf -ne 'y') {
        write-error "Execution has been aborted."
        Exit
    }
}


Function CreateTableFromDistList
{
    Param ($distListName)
    
    #Open Outlook and access Contacts folder
    $outlook = new-object -com Outlook.Application -ea 1
    $ns = $outlook.GetNamespace("MAPI")
    $contacts = $outlook.session.GetDefaultFolder(10)

    #Attempt to open distribution list (if existing)
    $dist = $contacts.Items | where {$_.DLName -eq $distListName}


    #http://blogs.msdn.com/b/rkramesh/archive/2012/02/02/creating-table-using-powershell.aspx

    #Start table
    $tabName = "DistListTable"

    #Create Table object
    $table = New-Object system.Data.DataTable “$tabName”

    #Define Columns
    $col1 = New-Object system.Data.DataColumn Name,([string])
    $col2 = New-Object system.Data.DataColumn StudentId,([string])
    $col3 = New-Object system.Data.DataColumn Email,([string])
    $col4 = New-Object system.Data.DataColumn Login,([string])

    #Add the Columns
    $table.columns.add($col1)
    $table.columns.add($col2)
    $table.columns.add($col3)
    $table.columns.add($col4)


    #$regex = [RegEx]'(?i)\s[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\s'

    foreach($entry in $dist.OneOffMembers) {

        #Student ID
        $studentId = (Select-String -InputObject $entry -Pattern '\([0-9]*\)' -AllMatches).Matches
        #write-host $studentId
        #write-host $studentId.Count
        if ($studentId.Count -gt 0) {
        
            $entry = $entry -replace $studentId[0] -replace ""

            $studentId = $studentId -replace "\(" -replace ""
            $studentId = [String] $studentId -replace "\)" -replace ""
        }
        #write-host $studentId

        #E-mail
        $email = (Select-String -InputObject $entry -Pattern '\w+@\w+\.\w+(.\w+)+' -AllMatches).Matches

        #write-host $email
    
        $entry = $entry -replace $email -replace ""

        #Name
        $name = (Select-String -InputObject $entry -Pattern '[A-Za-z]+( [A-Za-z]+)+' -AllMatches).Matches

        #write-host $name

        #Add data to table
    
        #Create a row
        $row = $table.NewRow()

        #Enter data in the row
        $row.Name = $name[0]
        $row.StudentId = $studentId
        $row.Email = $email[0]
        $row.Login = $row.Email.Split("@")[0] # extract login

        #Add the row to the table
        $table.Rows.Add($row)
    }

    return $table
}

Function RestartVM
{
    # VM to be renamed, maximum number of retries when restarting VM, indicates whether VM is started after stopping
    Param($myVmToBeRenamed, $maxRetries, $startAfterStopping)

    #Start, stop, and start, so the naming takes effect.
    $myVmToBeRenamed | Start-CIVM -ErrorAction Stop

    write-host "Starting VM $myVmToBeRenamed."

    $ct = 0;

    while($ct -le $maxRetries) {

        Start-Sleep -Seconds 5

        $err = $null;
        $ct = $ct + 1;
        write-host "Attempting to stop VM $myVmToBeRenamed ..."
        write-host "Try $ct"
        $myVmToBeRenamed | Stop-CIVMGuest -Confirm:$false -ErrorAction SilentlyContinue -ErrorVariable err
        if($err -ne $null)
        {
            write-host "Error: $err"
            if($ct -gt $maxRetries)
            {
                if($promptForFailedRestarts){
                    $conf = Read-Host "Stopping of VM $myVmToBeRenamed failed. Have reached $maxRetries retries). Want to continue ('Yes', 'Y' or 'y')?"
                    if($conf -eq 'Y' -or $conf -eq 'Yes' -or $conf -eq 'y') 
                    {
                        Write-Host "Continuing for three more iterations."
                        $maxRetries = $maxRetries + 3
                    } else {
                        Write-Error "Aborting execution after failed restarts."
                        $failedVMRestarts += $myVmToBeRenamed
                        #Aborting here
                    }
                } else {
                    Write-Error "Aborting execution after failed restarts."
                    $failedVMRestarts += $myVmToBeRenamed
                    #Aborting here
                }
            }
            else
            {
                Write-Host "Stopping of VM $myVmToBeRenamed failed. Will try again (Retry $ct, max: $maxRetries)."
            }
        } else {
            write-host "Stopping successful."
            $ct = $maxRetries + 1;
        }
    }

    if($startAfterStopping)
    {
        $myVmToBeRenamed | Start-CIVM -ErrorAction Stop

        write-host "Restart completed."
    }

}

#Use that function to restart all failed VMs
Function RestartVMs
{
    Param($vmsToBeRestarted, $retries)
    $vmsToBeRestarted | %{ RestartVM $_ $retries $true }
}

Function StopVM
{
    Param($myVmToBeRenamed, $maxRetries)
    RestartVM $myVmToBeRenamed $maxRetries $false
}

Function SetUpVApp
{
    Param ($srcTemplate, $vAppName, $user, $setPermission, $permissionLevel, $startVApp, $customizeHostnames, $VmNamesToChange)

    write-host "Creating $vAppName from template $srcTemplate (VDC $vdc) for user $user with permission $permissionLevel (Starting vApp: $startVApp, renaming hosts: $customizeHostnames, restart of new VMs (whether or not customized): $restartVMsAfterVAppStart, restart of existing VMs: $restartVMsEvenIfExisting)."

    #Retrieve the source vApp template for your new vApp - and stop if it does not exist

    $myVAppTemplate = Get-CIVAppTemplate -Name $srcTemplate -Catalog $catalog -ErrorAction Stop

    # Check first whether it exists
	if (Get-CIVApp -Name $vAppName -ErrorAction SilentlyContinue) {

        write-host "vApp $vAppName already exists, creation aborted." 
        if($setPermission) {
            write-host "Refining permissions. Does it affect the runtime status (start/stop/restart)? $restartVMsEvenIfExisting"
        }
        
        # Prevent start
        $startVApp = $false
    } else {

        # Create new vApp.
        $myVApp = New-CIVApp -Name $vAppName -VAppTemplate $myVAppTemplate -OrgVdc $myOrgVDC

        write-host "vApp $vAppName created for user $user."
    }

    if ($customizeHostnames) {

        # Assign vApp
        $myVApp = Get-CIVApp -Name $vAppName -OrgVdc $myOrgVDC -ErrorAction Stop
        
        $VmNamesToChange.Keys | ForEach-Object {

            $targetHostName = $VmNamesToChange.Item($_)

            # Do the customization for each vm name entry
            $myVmToBeRenamed = $myVApp | Get-CIVM -Name $_ -ErrorAction Stop

            $custom = $myVmToBeRenamed.ExtensionData.getGuestCustomizationSection()

            $custom.Enabled = $true

            # Allow user-specific customisation
            if ($appendUserToCustomHostname) {
                $targetHostName = $targetHostName + '-' + $user
            }

            write-host "Renaming VM $_'s hostname to $targetHostName"

            $custom.ComputerName = $targetHostName

            $custom.AdminAutoLogonCount = 0

            $custom.UpdateServerData()

            write-host "Restarting VM ..."

            # Restart VM, so effect takes place (max. 10 attempts)
            RestartVM $myVmToBeRenamed 10 $true
        }
    }

    if ($setPermission) {
        # Check for existence first
        $ct = Get-CIUser | where {$_.Name -eq $user} | measure

        # Only continue if user exists
        if($ct.Count -eq 1) {
            #Deletes existing ACL
            Get-CIAccessControlRule -Entity $vAppName -User $user | Remove-CIAccessControlRule -Confirm:$false -ErrorAction Stop

            if ($permissionLevel -ne 'None') {
                # Only reassign ACL if it is not 'None'
                New-CIAccessControlRule -Entity $vAppName -User $user -AccessLevel $permissionLevel -Confirm:$false
            }

            write-host "Permissions set to $permissionLevel (vApp: $vAppName, User: $user)"
        } else {
            Write-Error "User $user not found. Permissions NOT SET."
        }
    }

    # Marker whether VMs have been restarted independent of customization (to avoid double restart of VMs)
    $customizingIndependentVMRestart = $false

    # No need to start vApps if renaming - renaming requires restart
    if (!$customizeHostnames -and $startVApp) {
        Get-CIVApp -Name $vAppName -OrgVdc $myOrgVDC | Start-CIVApp

        write-host "VApp $vAppName started."

        # Restart all VMs after vApp start
        if($restartVMsAfterVAppStart) {
            write-host "Attempting to restart all VMs contained in new vApp."
            $myVApp = Get-CIVApp -Name $vAppName -ErrorAction Stop
            $vmsToRestart = $myVApp | Get-CIVM -ErrorAction Stop
            RestartVMs $vmsToRestart 10
            # Avoid repeated restarts independent of creation and customizing
            $customizingIndependentVMRestart = $true
        }
    }

    # Explicit restart - independent of customizing, but not done if just restarted after creation
    if($restartVMsEvenIfExisting -and !$customizingIndependentVMRestart) {
        write-host "Attempting to restart all VMs contained in existing vApp $vAppName."
        $myVApp = Get-CIVApp -Name $vAppName -OrgVdc $myOrgVDC -ErrorAction Stop
        $vmsToRestart = $myVApp | Get-CIVM -ErrorAction Stop
        RestartVMs $vmsToRestart 10
    }

}


function DeleteVApp
{
    Param($appName)

    $vApp = (Get-CIVApp -Name $appName -ErrorAction SilentlyContinue)

    if($vApp) {
        Stop-CIVApp $vApp -Confirm:$false -ErrorAction SilentlyContinue
        Remove-CIVApp $vApp -Confirm:$false
        write-host "VApp $appName deleted."
    } else {
        write-host "VApp $appName could not be found (and thus not be deleted)."
    }

}

# Checks for existence of vCloud users based on table (column 'Login') as produced by CreateTableFromDistList
function CheckForVCloudUsers
{
    Param ($table)
    #Check logins in vCloud
    $logCount = 0
    $notRegistered = @()

    $table | %{
        $name = $_["Login"]
        write-host "$logCount : Checking user $name"

        $ct = Get-CIUser | where {$_.Name -eq $name} | measure

        if($ct.Count -eq 0) {
            write-host "$name is not registered."
            $notRegistered += $name
        }
        $logCount = $logCount + 1
    }
    # Return the non-registered ones
    return $notRegistered
}


#### EXECUTION starts here ####

# If you just want to read the functions and settings without performing any action on vCloud
#Exit

$startTime = Get-Date
write-host "Start: $($startTime.ToString('u'))"

if ($readDistListToTable) {
    $table = CreateTableFromDistList $distListName
}

# If you just want to read the users from Outlook without performing any action on vCloud
#Exit

if ($testUserExistence) {
    $nonRegistered = CheckForVCloudUsers $table
}

if ($createVAppsForUsers -or $setPermissions -or $customizeHostnames) {

    #Select specific entry (e.g. for debugging): $table[index] (e.g. $table[1])
    $table | % {
        $name = $_["Login"]
        $appName = $vAppPrefix + $name + $vAppSuffix

        write-host "=== Processing CREATION or MODIFICATION of vApp $appName. ==="

        SetUpVApp $vAppTemplate $appName $name $setPermissions $permission $startVApps $customizeHostnames $VmNamesToChange

        write-host "=== CREATION or MODIFICATION of vApp $appName completed. ==="
    }
    write-host "***** CREATION or MODIFICATION OF vAPPS COMPLETED *****"
}

if ($stopVApps) {
    
    $table | % {
        $name = $_["Login"]
        $appName = $vAppPrefix + $name + $vAppSuffix

        write-host "=== Stopping vApp $appName. ==="

        $vApp = (Get-CIVApp -Name $appName -ErrorAction SilentlyContinue)

        if($vApp -and $vApp.Status -eq 'PoweredOn') {
            write-host " Found running vApp $appName. Initialising shutdown."
            Stop-CIVApp $vApp -Confirm:$false
        } else {
            if($vApp) {
                write-host " VApp $appName is not powered on."
            } else {
                write-host " VApp $appName does not exist."
            }
        }
        write-host "=== Stopping vApp $appName completed. ==="
    }
    write-host "***** STOPPING OF vAPPS COMPLETED *****"
}

if ($deleteVAppsForUsers) {
    
    $table | % {
        $name = $_["Login"]
        $appName = $vAppPrefix + $name + $vAppSuffix

        write-host "=== Processing DELETION of vApp $appName. ==="

        DeleteVApp $appName

        write-host "=== DELETION of vApp $appName completed. ==="
    }
    write-host "***** DELETION OF vAPPS COMPLETED *****"
}

$endTime = Get-Date
Write-Host "End: $($endTime.ToString('u'))"

Write-Host "Duration: $( New-TimeSpan $startTime $endTime)"


#Test for failed VM restarts
while($failedVMRestarts.Length -gt 0){

    $conf = Read-Host "Restart of $failedVMRestarts.Length VMs failed. Do you want to attempt their restart now ('Yes', 'Y' or 'y')?"
    if($conf -eq 'Y' -or $conf -eq 'Yes' -or $conf -eq 'y') {
        Write-Host "Attempting to restart failed VMs."
        $oldFailedRestarts = $failedVMRestarts
        $failedVMRestarts = @()

        $startTime = Get-Date
        Write-Host "Start: $($startTime.ToString('u'))"

        RestartVMs $oldFailedRestarts 10

        $endTime = Get-Date
        Write-Host "End: $($endTime.ToString('u'))"
        Write-Host "Duration: $( New-TimeSpan $startTime $endTime)"
    }
    Write-Host "Finished restart attempts of failed VMs."
}

Write-Host "All processing finished."