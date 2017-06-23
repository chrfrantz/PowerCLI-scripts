<#

    Purpose: This script generates a Outlook distribution list from a given CSV file (requires at least first name, last name and e-mail address).

    Instructions:
    - Start PowerShell ISE
    - Activate script execution in PowerShell ISE (Command: Set-ExecutionPolicy RemoteSigned)
    - Start Outlook
    - Specify desired distribution list name
    - Choose CSV file name
    - Set column headers correctly
    - Ensure that your shell is in the same directory the CSV file is located in
    - Run script
    - Deactivate script execution (security risk): Set-ExecutionPolicy Restricted
    
    Author: Christopher Frantz
    URL: https://github.com/chrfrantz/PowerShell-scripts.git
#>

#Distribution list name in Outlook to be created
$distListName = 'OutlookDistributionList'

#File path to csv file with columns produced by EBS: FORENAME, SURNAME, email11
#produced by EBS Report 'Class List with Email Addresses'
$FilePath = 'OP_Class List Email.csv'

#CSV table fields
$FirstName = 'FORENAME'
$LastName = 'SURNAME'
$Email = 'email11'

#Resources

#good overview on constants for Outlook folders
#http://blogs.msdn.com/b/jmanning/archive/2007/01/25/using-powershell-for-outlook-automation.aspx
#reference to use of createRecipient
#http://stackoverflow.com/questions/375148/add-new-records-to-private-outlook-distribution-list
#Generic scripts for full import of contact properties
#https://dl.dropboxusercontent.com/u/62204506/Blog/Import-OutlookContacts.txt
#Accessing properties by name
#http://stackoverflow.com/questions/27642169/looping-through-each-noteproperty-in-a-custom-object


[string]$Delimiter = ','

#Stop execution if anything goes wrong
$ErrorActionPreference = "Stop"

#Test for CSV file first before continuing
Try {
    $csv = Import-Csv $FilePath -Delimiter $Delimiter
} Catch {
    write-error "CSV file not found."
}

#Open Outlook and access Contacts folder
$outlook = new-object -com Outlook.Application -ea 1
$ns = $outlook.GetNamespace("MAPI")
$contacts = $outlook.session.GetDefaultFolder(10)

#Attempt to open distribution list (if existing)
$dist = $contacts.Items | where {$_.DLName -eq $distListName}

#If distribution list does not exist, create it
If(!$dist){

    #Create new distribution list
    $dist = $contacts.Items.Add(7)
    $dist.DLName = $distListName

    Write-host "Created new distribution list $distListName"
}

#Add individual entries from CSV
foreach($entry in $csv){
    $newMember = $ns.CreateRecipient($entry."$($FirstName)" + $entry."$($LastName)" + '<' + $entry."$($Email)" + '>')
    $newMember.resolve()
    $dist.AddMember($newMember)
    Write-host "Added" $entry."$($FirstName)" $entry."$($LastName)" " (" $entry."$($Email)" ")"
}

#Finally save distribution list
$dist.save()