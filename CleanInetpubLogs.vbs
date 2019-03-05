'////////////////////////////////////////////////////////////////////////////////////
'/
'/  Author:         Lucas Halbert <https://www.lhalbert.xyz>
'/  Date:           12.15.2017
'/  Last Edited:    03.15.2018
'/  Version:        2018.03.15
'/  Purpose:        Cleans logs in the inetpub directory 
'/  Description:    
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
'/  Revisions:  03.15.2018  Add options and comments to script
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
Const strTempDir = "c:\Temp"                              '/ String
strVerboseLogging=False                                   '/ Boolean
strLogFile = strTempDir & "\inetpublog_automation.log"    '/ String


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
If Not objFSO.FileExists(strLogFile) Then
    Set objLog = objFSO.CreateTextFile(strLogFile,True)
    objLog.Close
End If

'/  Open Log file
Set objLog = objFSO.OpenTextFile(strLogFile,ForAppending,True)


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
'/  Subroutine to display script usage
'/
'////////////////////////////////////////////////////////////////////////////////////
Sub DisplayUsage()
    Call WriteLogData("Script expects either two or three arguments", True) 
    Call WriteLogData("cscript CleanInetpubLogs.vbs <inetpub-log-path> <max-log-age[in days]> <verbose logging>", True)
    Call WriteLogData("cscript CleanInetpublogs.vbs C:\inetpub\logs\LogFiles 30 [True/False]", True)
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
'/  Main Script
'/
'////////////////////////////////////////////////////////////////////////////////////
Call WriteLogData(vbCrLf & vbCrLf & _
             "/**********************************/" & vbCrLf & _
             "/*  Starting inetpub log cleanup  */" & vbCrLf & _
             "/**********************************/" & vbCrLf, False)

'/ Parse and collect Command line Arguments
If (WScript.Arguments.Count < 2) Then
    Call WriteLogData(Now & ": ERROR: Too few arguments provided. Exiting...", True)
    Call DisplayUsage()
    AppExit(1)
ElseIf (WScript.Arguments.Count = 2) Then
    strLogPath=WScript.Arguments(0)
    intMaxAge=WScript.Arguments(1)
ElseIf (WScript.Arguments.Count = 3) Then
    strLogPath=WScript.Arguments(0)
    intMaxAge=WScript.Arguments(1)
    strVerboseLogging=WScript.Arguments(2)
Else
    Call WriteLogData(Now & ": ERROR: Too many arguments provided. Exiting...", True)
    CCall DisplayUsage()
    AppExit(1)
End If






'/ Set log file directrion
strLogFolder = "c:\inetpub\logs\LogFiles"

'/ Set Max age of log files in days
intMaxAge = 30



Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colFolder = objFSO.GetFolder(sLogFolder)

WScript.StdOut.Write("Checking logs in " & sLogFolder & vbCrLf)
WScript.StdOut.Write("Folder: " & colFolder.Path & vbCrLf)

For Each colSubfolder in colFolder.SubFolders
    Set objFolder = objFSO.GetFolder(colSubfolder.Path)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        iFileAge = now-objFile.DateCreated
        if iFileAge > (iMaxAge+1)  then
            WScript.StdOut.Write(objFile.Name & " is " & iFileAge & " days old and will be deleted" & vbCrLf)
            objFSO.deletefile objFile, True
        else
            WScript.StdOut.Write(objFile.Name & " is " & iFileAge & " days old and will NOT be deleted" & vbCrLf)
        end if
    Next
Next



Set objFSO = Nothing