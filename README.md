# Install_WMF-5.1

### Install Windows Management Framework v5.1 

Installs the Windows Management Framework version 5.1. Checks that the .NET Framework dependecies are met and installs the appropriate .NET redistributable if needed. This script takes a single argument which is a CIFS/SMB share path. This script requires that both the .NET Framework redistributable installer and the various Windows Management Framework installers are located in the command line specified share path and trusted by the network (ie: unblock the internet downloaded files).

#### Revisions:
1. <b>11.10.2017</b>  Fix IsPowerShellInstalled function to account for PowerShell versions lower than 3.x. Set Constant for variables that don't change.
2. <b>11.09.2017</b>  Change GetDotNetInformation function to GetMaxInstalledDotNetVersion to account for multiple versions of .NET installations reported out of order. Troubleshoot untrusted internet downloaded installers. Fix looping mechanism in OS and .NET information collection functions.
3. <b>11.07.2017</b>  Write logging and terminal output functions with verbose output option. Write function to check if the Windows  Update Service is running. Write GetDotNetInformation and IsDotNetFrameworkInstalled functions to ensure that  .NET Framework dependencies are met. Write functions to install .NET framework dependencies if they are missing. Write clean update function to destroy remaining open objects.
4. <b>11.03.2017</b>  Write functionality to pull installer from network share. Write functions to perform actual install of management framework
5. <b>10.13.2017</b>  Write function to check the OS architecture. Write the SelectInstaller function and insert necessary download URLs.
6. <b>10.12.2017</b>  Configure logging, Write function to gather necessary OS information. 
7. <b>10.05.2017</b>  Initial Draft