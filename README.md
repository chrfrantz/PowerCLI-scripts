# PowerShell-scripts
Scripts to automate administrative tasks

## CsvOutlookDistListConverter

### CsvToOutlookDistList.ps1

* Reads from CSV file including first name, last name and e-mail and creates Outlook distribution list.

### OutlookDistListToCsv.ps1

* Reads from Outlook distribution list and generates CSV.

## PowerCLI

All the following scripts rely on [PowerCLI](https://code.vmware.com/tool/vsphere_powercli/6.0), which needs to be installed on the client machine.

### ManageVApps

#### TemplateGlobals.ps1

* Contains all global variables for vSphere servers, organisation and VCD.

#### ConnectToVSphere.ps1

* Change `username` in script to administrative user.
* Run this script to connect to vSphere before running ManageVAppsForDistListMembers.ps1.

#### ManageVAppsForDistListMembers.ps1

* Check variables in script and adapt to your need. Most relevant ones:

  * `$distListName` Outlook distribution list name for which vApps should be instantiated. (User identifiers are extracted from the e-mail addresses in the distribution list.)

  * `$vAppTemplate` vApp template to use for instantiation.

  * `$vAppPrefix` Prefix for each vApp name

  * `$vAppSuffix` Suffix for any vApp

  * `$permission` Permission assigned to user (based on user identifier extracted from e-mail address)

* Check remaining variables.
* Run script.

### ManageVirtualNetworks

#### SetPromiscuousModeOnVSphere.ps1

* Change `username` in script to administrative user.
* Change server (`$server`) to match target server.
* Modify network name pattern (`$networkNameToModify`) to match name of virtual networks for which promiscuous mode and forged transmits should be activated.
* Run script.
