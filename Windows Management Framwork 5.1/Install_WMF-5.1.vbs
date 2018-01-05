'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Author:         Lucas Halbert <https://www.lhalbert.xyz>
'/  Date:           10.05.2017
'/  Last Edited:    01.05.2018
'/  Purpose:        Install Windows Management Framework 5.1
'/  Description:    Installs the Windows Management Framework version 5.1. Checks
'/                  that the .NET Framework dependecies are met and installs the
'/                  appropriate .NET redistributable if needed. This script takes a
'/                  single argument which is a CIFS/SMB share path. This script
'/                  requires that both the .NET Framework redistributable installer
'/                  and the various Windows Management Framework installers are
'/                  located in the command line specified share path and trusted by
'/                  the network (ie: unblock the internet downloaded files).
'/  License:        BSD 3-Clause License
'/
'/  Copyright (c) 2017, Lucas Halbert
'/  All rights reserved.
'/  
'/  Redistribution and use in source and binary forms, with or without
'/  modification, are permitted provided that the following conditions are met:
'/  
'/  * Redistributions of source code must retain the above copyright notice, this
'/    list of conditions and the following disclaimer.
'/  
'/  * Redistributions in binary form must reproduce the above copyright notice,
'/    this list of conditions and the following disclaimer in the documentation
'/    and/or other materials provided with the distribution.
'/  
'/  * Neither the name of the copyright holder nor the names of its
'/    contributors may be used to endorse or promote products derived from
'/    this software without specific prior written permission.
'/  
'/  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'/  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'/  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'/  DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
'/  FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
'/  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'/  SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
'/  CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
'/  OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'/  OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'/
'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Revisions:  01/05/2018  Add recursive call to start the Windows Update service
'/                          and to check if it properly started. Add Sleep subrouteine
'/                          to make waiting simpler. Fix .NET Framework detection.
'/
'/              12.18.2017  Add logging to the wusa MSU install.
'/
'/              11.10.2017  Fix IsPowerShellInstalled function to account for
'/                          PowerShell versions lower than 3.x. Set Constant for
'/                          variables that don't change.
'/
'/              11.09.2017  Change GetDotNetInformation function to 
'/                          GetMaxInstalledDotNetVersion to account for multiple
'/                          versions of .NET installations reported out of order.
'/                          Troubleshoot untrusted internet downloaded installers.
'/                          Fix looping mechanism in OS and .NET information
'/                          collection functions.
'/
'/              11.07.2017  Write logging and terminal output functions with verbose
'/                          output option. Write function to check if the Windows 
'/                          Update Service is running. Write GetDotNetInformation
'/                          and IsDotNetFrameworkInstalled functions to ensure that 
'/                          .NET Framework dependencies are met. Write functions to
'/                          install .NET framework dependencies if they are missing.
'/                          Write clean update function to destroy remaining open
'/                          objects.
'/
'/              11.03.2017  Write functionality to pull installer from network
'/                          share. Write functions to perform actual install of
'/                          management framework
'/
'/              10.13.2017  Write function to check the OS architecture. Write the
'/                          SelectInstaller function and insert necessary download
'/                          URLs.
'/
'/              10.12.2017  Configure logging, Write function to gather necessary OS
'/                          information. 
'/
'/              10.05.2017  Initial Draft
'/
'////////////////////////////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Initialize Variables
'/
'////////////////////////////////////////////////////////////////////////////////////
'/  Declare Variables
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1 
Const ForWriting = 2
Const ForAppending = 8
Const strComputer = "."                                   '/ String
Const strDotNETSourceFileName="NDP452-KB2901954-Web.exe"  '/ String
Const strExpectedDotNETVersion="4.5.2"                    '/ String
Const strExpectedPSVersion="5.1.14409.1012"               '/ String
Const strTempDir = "c:\Temp"                              '/ String
strVerboseLogging=False                                   '/ Boolean
strWMFLog = strTempDir & "\WMF_Installer.log"       '/ String
strDotNETLog = strTempDir & "\.NET_Installer.log"   '/ String
strWUSALog = strTempDir & "\WUSA.log"               '/ String
strOsVersion=""                                     '/ Float
strOsProductType=""                                 '/ Int
strOsArchitecture=""                                '/ String
strUrl=""                                           '/ String
strMaxInstalledDotNETVersion=""                     '/ String
strInstalledPSVersion=""                            '/ String


'/  Access necessary interfaces
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Configure logging to file
'/
'////////////////////////////////////////////////////////////////////////////////////
'/  Create Temp directory if it doesn't exist
If Not objFSO.FolderExists(strTempDir) Then
    Set objFldr = objFSO.CreateFolder(strTempDir)
End If

'/  Create Log File if it doesn't exist
If Not objFSO.FileExists(strWMFLog) Then
    Set objLog = objFSO.CreateTextFile(strWMFLog,True)
    objLog.Close
End If

'/  Open Log file
Set objLog = objFSO.OpenTextFile(strWMFLog,ForAppending,True)


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to exit and cleanup
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub AppExit(exitCode)
    Call WriteLogData(Now & ": INFO: Exiting app with exit code (" & exitCode & ")", False)
    Set objNetwork = Nothing
    Set objFSO = Nothing
    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objSysInfo=Nothing
    Set Shell = Nothing
    Wscript.Quit(exitCode)
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to write log data with verbose option
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub WriteLogData(str, verbose)
    WriteToLog(str)
    if strVerboseLogging OR verbose Then
        WriteToTerminal(str)
    End If
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to write string to terminal
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub WriteToTerminal(str)
    WScript.StdOut.WriteLine(str)
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to write string to log file
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub WriteToLog(str)
    objLog.WriteLine(str)
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to sleep for specified number of milliseconds
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub Sleep(time)
    WScript.Sleep(time)
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to Collect OS information
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub GetOsInformation()
    Call WriteLogData(Now & ": INFO: Collecting OS Information...", True)
    On Error Resume Next
    ' Query WMI
    Set objWMIService = GetObject("winmgmts://" & strComputer & "/root/cimv2")
    ' Get the error number 
    If Err.Number Then
        strMsg = vbCrLf & strComputer & vbCrLf & _
                 "Error # " & Err.Number & vbCrLf & _
                 Err.Description & vbCrLf & vbCrLf
        
        Call WriteLogData(Now & ": " & strMsg, True)
        AppExit(1)
    End If

    ' Collect OS information 
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
    For Each objItem in colItems
        strMsg = vbCrLf & "/******** OS Information ********/" & vbCrLf & _
                 "Computer Name   : " & objItem.CSName & vbCrLf & _
                 "Windows Version : " & objItem.Version & vbCrLf & _
                 "ServicePack     : " & objItem.CSDVersion & vbCrLf & _
                 "Product Type    : " & objItem.ProductType & vbCrLf & _
                 "OS Architecture : " & objItem.OSArchitecture & vbCrLf & _
                 "/********************************/" & vbCrLf
    
    ' Get the first two digits from the version string
    versionArray = Split(objItem.Version, ".", -1, 1)
    strOsVersion=versionArray(0) & "." & versionArray(1)
    oSSubVersion=versionArray(2)
    strOsProductType=objItem.ProductType
    strOsArchitecture=objItem.OSArchitecture

    ' Cleanup
    Set objWMIService = Nothing
    Set colItems = Nothing
    Next


    Call WriteLogData(strMsg, False)
    Call WriteLogData(Now & ": INFO: OS Version: " & strOsVersion, False)
    Call WriteLogData(Now & ": INFO: OS Sub-Version: " & oSSubVersion, False)
    Call WriteLogData(Now & ": INFO: OS Product Type: " & strOsProductType, False)
    Call WriteLogData(Now & ": INFO: OS Architecture: " & strOsArchitecture, False)
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to Collect .NET Framework information
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub GetMaxInstalledDotNetVersion
    Call WriteLogData(Now & ": INFO: Collecting Installed .NET Framework Versions...", True)
    On Error Resume Next

    '/ Set impersonation for registry query
    set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

    '/ Get all keys within strKeyPath
    strValueName = "Version"
    strKeyPath = "SOFTWARE\Microsoft\NET Framework Setup\NDP"
    objReg.EnumKey HKEY_LOCAL_MACHINE,strKeyPath,allSub

    strMsg = vbCrLf & "/** .NET Framework Information **/" & vbCrLf

    For Each strKey In allSub
        objReg.EnumKey HKEY_LOCAL_MACHINE,strKeyPath & "\" & strKey, allSubTwo
        For Each strKeyTwo In allSubTwo
            objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath & "\" & strKey & "\" & strKeyTwo,strValueName,strValue
            if strValue <> Empty Then
                Call WriteLogData(Now & ": INFO: Detected .NET Version: " & strValue, False)
                strMsg = strMsg & _
                    ".NET Version    : " & strValue & vbCrLf

                If strValue > strMaxInstalledDotNETVersion Then
                    strMaxInstalledDotNETVersion = strValue
                End If
            End If
        Next
    Next
    strMsg = strMsg & _
        "/********************************/" & vbCrLf

    Set objReg = Nothing

    Call WriteLogData(strMsg, False)
    Call WriteLogData(Now & ": INFO: Max Installed .NET Version: " & strMaxInstalledDotNETVersion, False)



'/    Call WriteLogData(Now & ": INFO: Collecting Installed .NET Framework Versions...", True)
'/    On Error Resume Next
'/    
'/    ' Query WMI
'/    Set objWMIService = GetObject("winmgmts://" & strComputer & "/root/cimv2")
'/    
'/    ' Get the error number 
'/    If Err.Number Then
'/        strMsg = vbCrLf & strComputer & vbCrLf & _
'/                 "Error # " & Err.Number & vbCrLf & _
'/                 Err.Description & vbCrLf & vbCrLf
'/        
'/        Call WriteLogData(Now & ": " & strMsg, True)
'/        AppExit(1)
'/    End If
'/
'/    '/ Collect .NET framework information
'/    Set colItems = objWMIService.ExecQuery("Select Name, Version from Win32_Product Where Name Like 'Microsoft .NET Framework%'")
'/    strMsg = "/** .NET Framework Information **/" & vbCrLf
'/    For Each objItem in colItems
'/        strMsg = strMsg & _
'/            "Framework Name  : " & objItem.Name & vbCrLf & _
'/            ".NET Version    : " & objItem.Version & vbCrLf
'/
'/        If objItem.Version > strMaxInstalledDotNETVersion Then
'/            strMaxInstalledDotNETVersion = objItem.Version
'/        End If
'/
'/    ' Cleanup
'/    Set objWMIService = Nothing
'/    Set colItems = Nothing
'/    Next
'/    
'/
'/    strMsg = strMsg & _
'/        "/********************************/" & vbCrLf
'/
'/    Call WriteLogData(strMsg, False)
'/    Call WriteLogData(Now & ": INFO: Max Installed .NET Version: " & strMaxInstalledDotNETVersion, False)
End Sub

'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to install a standard windows executable 
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub InstallExe(installerPath)
    '/ Install EXE
    cmdLine =  installerPath & " /q /norestart /log:" & Chr(34) & strTempDir & strDotNETLog & Chr(34)

    Call WriteLogData(Now & ": INFO: Running " & cmdLine, False)

    Set WshShell = WScript.CreateObject("WScript.Shell")
    Return = WshShell.Run(cmdLine, 0, true)

    ' cleanup
    Set WshShell= Nothing
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Subroutine to install MSU
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub InstallMsu(installerPath)
    '/ Install MSU
    cmdLine ="wusa.exe " & installerPath & " /quiet /norestart /log:" & strWUSALog
    
    Call WriteLogData(Now & ": INFO: Running " & cmdLine, False)
    
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Return = WshShell.Run(cmdLine, 0, true)
    
    '/ cleanup
    Set WshShell=Nothing
End Sub


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to Determine if OS is 64-bit or not
'/
'////////////////////////////////////////////////////////////////////////////////////
Function IsOS64Bit()
    Call WriteLogData(Now & ": INFO: Checking OS Architecture...", False)
    
    intArch = StrComp(strOsArchitecture,"64-bit",1)
    'Architecture is not available on certain OSs, so check the registry
    If (intArch = -1) Then
        set Shell = CreateObject("WScript.Shell") 
        Shell.RegRead "HKLM\Software\Microsoft\Windows\CurrentVersion\ProgramFilesDir (x86)" 
        
        If Err.Number <> 0 Then 
            Call WriteLogData(Now & ": INFO: OS is Not 64-bit: " & Err.Description, False)
            IsOS64Bit = False
            Exit Function
        else   
            Call WriteLogData(Now & ": INFO: OS is 64-bit", False)
            IsOS64Bit = True
            Exit Function
        End If

        Set Shell = Nothing
    else
        Call WriteLogData(Now & ": INFO: OS is 64-bit", False)
        IsOS64Bit = True
        Exit Function
    End If
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to Determine if .NET Framework version
'/
'////////////////////////////////////////////////////////////////////////////////////
Function IsDotNetFrameworkInstalled(expectedVersion)
    Call WriteLogData(Now & ": INFO: Checking if .NET Framework dependencies are met", True)

    if Mid(strMaxInstalledDotNETVersion,1,5) >= expectedVersion Then
        Call WriteLogData(Now & ": INFO: Expected .NET Framework Version (" & expectedVersion & ") or higher already installed... Installed .NET Version (" & strMaxInstalledDotNETVersion & ")", False)
        IsDotNetFrameworkInstalled = True
        Exit Function
    Else
        Call WriteLogData(Now & ": INFO: Expected .NET Framework Version (" & expectedVersion & ") or higher not installed... Installed .NET Version (" & strMaxInstalledDotNETVersion & ")", False)
        IsDotNetFrameworkInstalled = False
        Exit Function
    End If
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to Determine if the expected PowerShell version is installed
'/
'////////////////////////////////////////////////////////////////////////////////////
Function IsPowerShellInstalled(expectedVersion)
    Call WriteLogData(Now & ": INFO: Checking if PowerShell is installed", True)
    Call WriteLogData(Now & ": INFO: Checking for PowerShell Version 3.x and above", False)

    '/ Set impersonation for registry query
    set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
     
    strKeyPath = "SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine"
    strValueName = "PowerShellVersion"
    objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

    '/ Handle Powershell version 2.x and down
    If IsEmpty(strValue) or ISNull(strValue) Then
        Call WriteLogData(Now & ": INFO: Checking for PowerShell Version 2.x and below", False)
        
        strKeyPath = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine"
        strValueName = "PowerShellVersion"
        objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
        installedVersion=strValue
    Else
        installedVersion=strValue
    End If

    Set objReg = Nothing

    '/ Compare installed version of PowerShell to the expected version
    If (StrComp(Mid(installedVersion,1,3),Mid(expectedVersion,1,3),1) = 0) Then
        Call WriteLogData(Now & ": INFO: Expected PowerShell Version (" & Mid(expectedVersion,1,3) & ") already installed... Installed PowerShell Version (" & installedVersion & ")", False)
        
        IsPowerShellInstalled=True
        Exit Function
    Else
        Call WriteLogData(Now & ": INFO: Expected PowerShell Version (" & Mid(expectedVersion,1,3) & ") not installed... Installed PowerShell Version (" & installedVersion & ")", False)
        
        IsPowerShellInstalled=False
        Exit Function
    End If
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to Determine if Windows Update Service is Running
'/
'////////////////////////////////////////////////////////////////////////////////////
Function IsWindowsUpdateServiceRunning(counter)
    Call WriteLogData(Now & ": INFO: Checking if Windows Update Service is running", True)

    If counter >= 3 Then
        Call WriteLogData(Now & ": ERROR: Windows Update Service (wuauserv) is not running. Restart failed 3 times.", True)
        IsWindowsUpdateServiceRunning=False
        Exit Function
    End If

    '/ Set impersonation for WMI query
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
    '/strWMIQuery = "Select * from Win32_Service Where Name = 'wuauserv' and state='Running'"
    strWMIQuery = "Select * from Win32_Service Where Name = 'wuauserv' and StartMode='Auto'"

    For Each service In objWMIService.ExecQuery(strWMIQuery)
        If service.State <> "Running" Then
            Call WriteLogData(Now & ": INFO: Windows Update Service (wuauserv) is not running. Attempting to start", True)
            service.StartService
            counter = counter + 1
            Sleep(2000)

            '/ Recursively call IsWindowsUpdateServiceRunning to check if sending start command worked
            IsWindowsUpdateServiceRunning=IsWindowsUpdateServiceRunning(counter)
            Exit Function
        Else
            Call WriteLogData(Now & ": INFO: Windows Update Service (wuauserv) is running.", False)
            IsWindowsUpdateServiceRunning=True
            Exit Function
        End If
    Next
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to Select the correct installer based on 
'/
'////////////////////////////////////////////////////////////////////////////////////
Function SelectWMFInstaller()
    '///////////////////////////////////////////////////////////////
    '/ Source: http://technet.microsoft.com/en-us/library/cc947846(WS.10).aspx
    '/ProductType="1" -> Client operating systems
    '/ProductType="2" -> Domain controllers
    '/ProductType="3" -> Servers that are not domain controllers
    '/
    '/ Installer Files:
    '/ W2K12-KB3191565-x64.msu
    '/ Win7AndW2K8R2-KB3191566-x64.zip
    '/ Win7-KB3191566-x86.zip
    '/ Win8.1AndW2K12R2-KB3191564-x64.msu
    '/ Win8.1-KB3191564-x86.msu
    '///////////////////////////////////////////////////////////////

    ' Select the installer based on the OS Version and Product Type
    Select Case strOsVersion
        Case "10.0"
            '/ strOsVersion 10.0 = Windows-10 and Server-2016
            Call WriteLogData(Now & ": INFO: OS: Windows 10", False)
            Call WriteLogData(Now & ": INFO: Windows 10 ships with WMF-5.1", False)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(0)
         
        Case "6.3"
            '/ strOsVersion 6.3 = Windows-8.1 and Server-2012R2
            If (StrComp(strOsProductType,"1",1) =0 ) Then
                If(IsOS64Bit()) then
                    Call WriteLogData(Now & ": INFO: OS: Windows 8.1-x64", False)
                    url="http://download.windowsupdate.com/d/msdownload/update/software/updt/2017/03/windowsblue-kb3191564-x64_91d95a0ca035587d4c1babe491f51e06a1529843.msu"
                    sourceFileName="Win8.1AndW2K12R2-KB3191564-x64.msu"
                Else 
                    Call WriteLogData(Now & ": INFO: OS: Windows 8.1-x86", False)
                    url="http://download.windowsupdate.com/d/msdownload/update/software/updt/2017/03/windowsblue-kb3191564-x86_821ec3c54602311f44caa4831859eac6f1dd0350.msu"
                    sourceFileName="Win8.1-KB3191564-x86.msu"
                End if 
            Else
                If(IsOS64Bit()) then
                    Call WriteLogData(Now & ": INFO: OS: Server 2012R2-x64", False)
                    url="http://download.windowsupdate.com/d/msdownload/update/software/updt/2017/03/windowsblue-kb3191564-x64_91d95a0ca035587d4c1babe491f51e06a1529843.msu"
                    sourceFileName="Win8.1AndW2K12R2-KB3191564-x64.msu"
                Else
                    Call WriteLogData(Now & ": ERROR: There is no Server 2012R2-x86", True)
                    Call WriteLogData(Now & ": INFO: Exiting...", False)
                    AppExit(1)
                End if 
            End If

        Case "6.2"
            '/ strOsVersion 6.2 = Windows-8.0 and Server-2012
            If (StrComp(strOsProductType,"1",1) =0 ) Then
                If(IsOS64Bit()) then
                    Call WriteLogData(Now & ": INFO: OS: Windows 8.0-x64", False)
                    Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-8.0x64", True)
                    Call WriteLogData(Now & ": INFO: Exiting...", False)
                    AppExit(1)
                Else 
                    Call WriteLogData(Now & ": INFO: OS: Windows 8.0-x86", False)
                    Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-8.0x86", True)
                    Call WriteLogData(Now & ": INFO: Exiting...", False)
                    AppExit(1)
                End if 
            Else
                If(IsOS64Bit()) then                
                    Call WriteLogData(Now & ": INFO: OS: Server 2012-x64", False)
                    
                    url="http://download.windowsupdate.com/d/msdownload/update/software/updt/2017/03/windows8-rt-kb3191565-x64_b346e79d308af9105de0f5842d462d4f9dbc7f5a.msu"
                    sourceFileName="W2K12-KB3191565-x64.msu"
                Else
                    Call WriteLogData(Now & ": ERROR: There is no Server 2012-x86", True)
                    Call WriteLogData(Now & ": INFO: Exiting...", False)
                    AppExit(1)
                End If
            End If

        Case "6.1"
            '/ strOsVersion 6.2 = Windows-7 and Server-2008R2            
            If (StrComp(strOsProductType,"1",1) =0 ) Then
                If(IsOS64Bit()) then
                    Call WriteLogData(Now & ": INFO: OS: Windows 7x64", False)
                    Call WriteLogData(Now & ": WARN: There is no direct download for the Windows-7/Server-2008R2 msu package", False)
                    
                    url=""
                    sourceFileName="Win7AndW2K8R2-KB3191566-x64.msu"
                Else
                    Call WriteLogData(Now & ": INFO: OS: Windows 7x86", False)
                    Call WriteLogData(Now & ": WARN: There is no direct download for the Windows-7/Server-2008R2 msu package", False)
                    
                    url=""
                    sourceFileName = "Win7-KB3191566-x86.msu"
                End if 
            Else
                If(IsOS64Bit()) then
                    Call WriteLogData(Now & ": INFO: OS: Server 2008R2", False)
                    Call WriteLogData(Now & ": WARN: There is no direct download for the Windows-7/Server-2008R2 msu package", False)
                    
                    url=""
                    sourceFileName="Win7AndW2K8R2-KB3191566-x64.msu"
                Else
                    Call WriteLogData(Now & ": ERROR: There is no Server 2008R2-x86", True)
                    Call WriteLogData(Now & ": INFO: Exiting...", False)
                    AppExit(1)
                End if 
            End If

        Case "6.0"
            '/ strOsVersion 6.0 = Windows-Vista and Server-2008
            Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-Vista, or Server-2008", True)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(1)

        Case "5.2"
            '/ strOsVersion 5.2 = Windows-XPx64, Server-2003, and Server-2003R2
            Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-XPx64, Server-2003, or Server-2003R2", True)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(1)
                       
        Case "5.1"
            '/ strOsVersion 5.1 = Windows-XP
            Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-XPx86, Server-2003, or Server-2003R2", True)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(1)

        Case "5.0"
            '/ strOsVersion 5.0 = Windows 2000
            Call WriteLogData(Now & ": ERROR: Window Management Framework V5.1 is not supported on Windows-2000", True)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(1)

        Case Else
            Call WriteLogData(Now & ": ERROR: Unknown OS Version (" & strOsVersion & ")", True)
            Call WriteLogData(Now & ": INFO: Exiting...", False)
            AppExit(1)
    End Select

    '/ Return source file name
    SelectWMFInstaller=sourceFileName
    Exit Function
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to copy installer from network share
'/
'////////////////////////////////////////////////////////////////////////////////////
Function CopyInstallerFromShare(file)
    Call WriteLogData(Now & ": INFO: Mapping drive to copy " & file & " from share", False)

    On Error Resume Next
    '/ Map network drive
    objNetwork.MapNetworkDrive "", strShare, False
    
    If Err.Number Then
        Call WriteLogData(Now & ": ERROR: Failed to add the share:" & strShare, True)
        Call WriteLogData(Now & ": ERROR: Number:" & Err.Number, True)
        Call WriteLogData(Now & ": ERROR: Description:" & Err.Description, True)
        Call WriteLogData(Now & ": ERROR: Exception" & Err.GetException(), True)
        AppExit(1)
    End If
 
    '/ Copy installer from share to temp directory
    sourcePath = strShare & "\" & file
    destinationPath = strTempDir & "\" & file

    Call WriteLogData(Now & ": INFO: Copying " & sourcePath & " to " & destinationPath, False)
    objFSO.CopyFile sourcePath , destinationPath , true

    If Err.Number Then
        Call WriteLogData(Now & ": ERROR: Failed to copy the WMF installer from the share:" & sourcePath, True)
        Call WriteLogData(Now & ": ERROR: Number:" & Err.Number, True)
        Call WriteLogData(Now & ": ERROR: Description:" & Err.Description, True)
        Call WriteLogData(Now & ": ERROR: Exception" & Err.GetException(), True)

        '/ Unmap network drive
        objNetwork.RemoveNetworkDrive strShare, True, False
        AppExit(1)
    End If
    
    '/ Unmap network drive
    objNetwork.RemoveNetworkDrive strShare, True, False
    
    If objFSO.FileExists(destinationPath) Then
        Call WriteLogData(Now & ": INFO: Copied " & file & " from:" & strShare & " to " & destinationPath, False)
        
        CopyInstallerFromShare=True
        Exit Function
    Else
        Call WriteLogData(Now & ": ERROR: Failed to copy " & file & " from:" & strShare & " to " & destinationPath, True)
        
        CopyInstallerFromShare=False
        Exit Function
    End If
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to install .NET Framework dependency
'/
'////////////////////////////////////////////////////////////////////////////////////
Function InstallDotNETFramework()
    '/ Copy Installer from share to temp directory
    If NOT (CopyInstallerFromShare(strDotNETSourceFileName)) Then
        InstallDotNETFramework=False
        Exit Function
    End if

    '/ Capture installer path
    installerPath = strTempDir & "\" & strDotNETSourceFileName
    
    '/ Capture the installer extension
    installerExtension=Right(installerPath,4)

    '/ Start the installation process
    Call WriteLogData(Now & ": INFO: Installing .NET Framework dependencies", True)
    If (StrComp(installerExtension,".msu",1) =0 ) Then
        Call WriteLogData(Now & ": INFO: " & strDotNETSourceFileName & " installer is an msu: Windows Update Standalone Installer", False)
        InstallMsu(installerPath)
    ElseIf (StrComp(installerExtension,".exe",1) =0 ) Then
        Call WriteLogData(Now & ": INFO: " & strDotNETSourceFileName &" installer is an exe: Standard Windows Executable", False)
        InstallExe(installerPath)
    Else
        Call WriteLogData(Now & ": ERROR: " & strDotNETSourceFileName & " Installer type unknown.", True)
        
        InstallDotNETFramework=False
        Exit Function
    End If
    InstallDotNETFramework=True
    Exit Function
End Function

'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Function to install Window Management Framework 5.1
'/
'////////////////////////////////////////////////////////////////////////////////////
Function InstallWMF()
    '/ Check if Windows update service is running
    If NOT IsWindowsUpdateServiceRunning(0) Then
        InstallWMF=False
        Exit Function
    End If

    '/ Select WMF installer for OS Version
    WMFInstallerFileName=SelectWMFInstaller()

    '/ Copy Installer from share to temp directory
    If NOT (CopyInstallerFromShare(WMFInstallerFileName)) Then
        InstallWMF=False
        Exit Function
    End If

    '/ Capture installer path
    installerPath = strTempDir & "\" & WMFInstallerFileName
    
    '/ Capture the installer extension
    installerExtension=Right(installerPath,4)
   
    '/ Start the installation process
    Call WriteLogData(Now & ": INFO: Installing Windows Management Framework", True)
    If ( StrComp(installerExtension,".msu",1) =0 ) Then
        Call WriteLogData(Now & ": INFO: "& WMFInstallerFileName & " installer is an msu: Windows Update Standalone Installer", False)
        InstallMsu(installerPath)
    ElseIf (StrComp(installerExtension,".exe",1) =0 ) Then
        Call WriteLogData(Now & ": INFO: " & WMFInstallerFileName &" installer is an exe: Standard Windows Executable", False)
        InstallExe(installerPath)
    Else
        Call WriteLogData(Now & ": ERROR: " & WMFInstallerFileName & " installer type unknown.", True)
        
        InstallWMF=False
        Exit Function
    End If
    
    '/ Check if a reboot is needed
    Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")

    Call WriteLogData(Now & ": INFO: Reboot required? " & objSysInfo.RebootRequired, False)
    If ( StrComp(objSysInfo.RebootRequired,"True",1) =0 ) Then
        Call WriteLogData(Now & ": INFO: System requires a reboot to fully upgrade the Windows Management Framework", True)
        
        Set objSysInfo=Nothing
        InstallWMF=True
        Exit Function
    End If
    Set objSysInfo=Nothing
    InstallWMF=True
    Exit Function
End Function


'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Main Script
'/
'////////////////////////////////////////////////////////////////////////////////////
Call WriteLogData(vbCrLf & vbCrLf & _
             "/********************************/" & vbCrLf & _
             "/*  Starting WMF-5.1 Installer  */" & vbCrLf & _
             "/********************************/" & vbCrLf, False)

'/ Parse Command line Arguments
If (WScript.Arguments.Count < 1) Then
    Call WriteLogData("Usage: (expects one or two arguments) <remote-sharepath> <verbose logging>", True)
    Call WriteLogData("For e.g. \\192.168.1.1\share [True/False]", True)
    Call WriteLogData(Now & ": ERROR: Too few arguments provided. Exiting...", True)
    AppExit(1)
End if

'/ Gather share name from command line arguments
If (WScript.Arguments.Count = 1) Then
    strShare=WScript.Arguments(0)
ElseIf (WScript.Arguments.Count = 2) Then
    strShare=WScript.Arguments(0)
    strVerboseLogging=WScript.Arguments(1)
End If


'/ Collect OS information
GetOsInformation


'/ Collect .NET information
GetMaxInstalledDotNetVersion


'/ Check if expected version of PowerShell is already installed
if (IsPowerShellInstalled(strExpectedPSVersion)) Then
    Call WriteLogData(Now & ": INFO: Windows Management Framework v" & strExpectedPSVersion & " is already installed. Exiting...", True)
    AppExit(0)
End If


'/ Check if expected version of .NET Framework is already installed
If NOT (IsDotNetFrameworkInstalled(strExpectedDotNETVersion)) Then
    Call WriteLogData(Now & ": INFO: .NET Framework Version " & strExpectedDotNETVersion & " is not yet installed.", False)
    If NOT (InstallDotNETFramework()) Then
        Call WriteLogData(Now & ": ERROR: A problem occured while installing .NET Framework" & strExpectedDotNETVersion & ". Exiting...", True)
        AppExit(1)
    End If
End If



If (InstallWMF()) Then
    Call WriteLogData(Now & ": INFO: Window Management Framework v" & strExpectedPSVersion & " has been successfully installed. Exiting...", True)
    AppExit(0)
Else
    Call WriteLogData(Now & ": ERROR: A problem occured while installing WMF-" & strExpectedPSVersion & ". Exiting...", True)
    AppExit(1)
End If