# PowerShell-scripts
Scripts to automate administrative tasks (Author: Christopher Frantz)

Most of those were developed to manage courses infrastructure at the Otago Polytechnic.

This file provides and overview over the folder structure and the contained scripts.

## CsvOutlookDistListConverter

Scripts related to the generation of Outlook distribution lists from CSV files and vice versa.

### CsvToOutlookDistList.ps1

* Reads from CSV file including first name, last name and e-mail and creates Outlook distribution list.

### OutlookDistListToCsv.ps1

* Reads from Outlook distribution list and generates CSV.

## PowerCLI

Scripts to automate routine tasks for vSphere and vCloud.

All the following scripts rely on [PowerCLI](https://www.vmware.com/support/developer/PowerCLI/) (version 6.0), which needs to be installed on the client machine. Ensure to include the features 'vSphere PowerCLI' and 'vCloud Air/vCD PowerCLI' (additional features are optional) during installation.

In PowerShell you will need to enable the execution of scripts (`Set-ExecutionPolicy RemoteSigned`).

### ManageVApps

#### TemplateGlobals.ps1

* Contains all global variables for vSphere servers, organisation and VCD.

#### ConnectToVSphere.ps1

* Change `username` in script to administrative username.
* Run this script (as administrator) to connect to vSphere before running ManageVAppsForDistListMembers.ps1.

#### ManageVAppsForDistListMembers.ps1

* Start Outlook (for distribution list reading).

* Start PowerShell in same security context as Outlook (e.g. both as administrator). If both applications run in different security contexts you will see a COM exception (CO_E_SERVER_EXEC_FAILURE).

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

### RetrieveIpAddresses

#### GetIpAddresses.ps1

* Follow instructions in script.