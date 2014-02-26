' Create symbolic link to MT4(build 600 and later) and MT5 terminal data path
' License: GNU General Public License Version 3
Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim arguments
Set arguments = WScript.Arguments

Dim sh
Set sh = CreateObject("WScript.Shell")

Dim currentDirectory
Dim appdataDirectory
If arguments.Count = 0 Then
    currentDirectory = fso.GetAbsolutePathName(".")
    appdataDirectory = sh.ExpandEnvironmentStrings("%APPDATA%")
Else
    currentDirectory = arguments(0)
    appdataDirectory = arguments(1)
End If

Dim wmi, os, value

Set wmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set os = wmi.ExecQuery("SELECT *FROM Win32_OperatingSystem")
For Each value in os
    If left(value.Version, 3) < 6.0 Then
        MsgBox "This script cannot be run on earlier to Windows Vista", 16
        WScript.Quit
    End If
Next

Dim app
Do While WScript.Arguments.Count = 0 and WScript.Version >= 5.7
    Set app = CreateObject("Shell.Application")
    app.ShellExecute "wscript.exe", """" & _
        WScript.ScriptFullName & """ """ & currentDirectory & _
        """ """ & appdataDirectory & """", "", "runas"

    WScript.Quit
loop

Function GetTerminalName(originTxt)
    Dim objOriginTxt, terminalName
    Set objOriginTxt = fso.OpenTextFile(originTxt, 1, False, -1)
    terminalName = objOriginTxt.ReadLine

    objOriginTxt.Close
    Set objOriginTxt = Nothing

    GetTerminalName = Right(terminalName, Len(terminalName) - InStrRev(terminalName, "\"))
End Function

Function CreateSymbolicLink(target, link)
    Dim cmd
    cmd = "cmd /c mklink /d """ & link & """ """ & target & """"

    sh.CurrentDirectory = currentDirectory
    sh.Run cmd, 0, true


    If fso.FolderExists(currentDirectory & "\" & link) Then
        CreateSymbolicLink = currentDirectory & "\" & link
    Else
        CreateSymbolicLink = ""
    End If
End Function

Dim terminalDataParentFolder
Set terminalDataParentFolder = fso.GetFolder(appdataDirectory & "\MetaQuotes\Terminal")

Dim subFolders
Set subFolders = terminalDataParentFolder.SubFolders

Dim f, path, originTxt, symlink
For Each f In subFolders
    path = terminalDataParentFolder.Path & "\" & f.name
    originTxt = path & "\origin.txt"
    If fso.FileExists(originTxt) Then
        symlink = CreateSymbolicLink(path, GetTerminalName(originTxt))
        If symlink = "" Then
            MsgBox "Failed to create symlink to " & path, 16
        End If
    End If
Next

MsgBox "Finished"
