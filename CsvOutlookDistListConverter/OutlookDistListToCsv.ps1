<#

    Purpose: This script generates a CSV file from an Outlook Distribution List.

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

#Distribution list name in Outlook
$distListName = 
'OutlookDistributionList'

#Output file
$FilePath = 'OP_Class List Email.csv'

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


#Display the table
$table | Export-Csv -Path $FilePath -NoTypeInformation
