' This file must stay pure ASCII. WScript is the console-free entry point.

Option Explicit

Dim shell
Dim fileSystem
Dim baseDirectory
Dim scriptPath
Dim powerShellPath
Dim command
Dim exitCode
Dim logDirectory

Set shell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

baseDirectory = fileSystem.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fileSystem.BuildPath(baseDirectory, "ReaderDataViewer.ps1")
powerShellPath = shell.ExpandEnvironmentStrings("%SystemRoot%") & _
    "\System32\WindowsPowerShell\v1.0\powershell.exe"
If Not fileSystem.FileExists(powerShellPath) Then
    powerShellPath = "powershell.exe"
End If

command = Chr(34) & powerShellPath & Chr(34) & _
    " -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File " & _
    Chr(34) & scriptPath & Chr(34)

shell.CurrentDirectory = baseDirectory
exitCode = shell.Run(command, 0, True)

If exitCode = 3 Then
    logDirectory = fileSystem.BuildPath( _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%"), _
        "ReaderDataViewer\logs")
    MsgBox _
        "Reader Data Viewer could not start." & vbCrLf & vbCrLf & _
        "The reason was written to the newest log file in:" & vbCrLf & _
        logDirectory, _
        vbExclamation, "Reader Data Viewer"
End If
