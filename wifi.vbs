Option Explicit

Dim shell, fso, outFile, profiles, profile, ssid, cmd, exec, line
Dim outputFilePath

outputFilePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & "\wifipassword.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Check for admin rights
If Not IsAdmin() Then
    ' Relaunch the script with admin rights
    RunAsAdmin
    WScript.Quit
End If

' Create or overwrite output file
Set outFile = fso.CreateTextFile(outputFilePath, True)

outFile.WriteLine "Retrieving saved Wi-Fi profiles and passwords..."
outFile.WriteLine

' Get profiles list
cmd = "netsh wlan show profiles"
Set exec = shell.Exec(cmd)

Do While Not exec.StdOut.AtEndOfStream
    line = Trim(exec.StdOut.ReadLine)
    If InStr(line, "All User Profile") > 0 Then
        ssid = Trim(Split(line, ":")(1))
        outFile.WriteLine "SSID: " & ssid

        ' Get key content for this profile
        cmd = "netsh wlan show profile name=""" & ssid & """ key=clear"
        Set exec = shell.Exec(cmd)

        Do While Not exec.StdOut.AtEndOfStream
            line = exec.StdOut.ReadLine
            If InStr(line, "Key Content") > 0 Then
                outFile.WriteLine Trim(line)
            End If
        Loop
        outFile.WriteLine
    End If
Loop

outFile.WriteLine "Done! Passwords saved in wifipassword.txt"
outFile.Close

MsgBox "Done! Passwords saved in wifipassword.txt", vbInformation, "Wi-Fi Password Retriever"

' Function to check if script is running as admin
Function IsAdmin()
    Dim wshShell, oExec, strUser
    Set wshShell = CreateObject("WScript.Shell")
    On Error Resume Next
    Set oExec = wshShell.Exec("net session >nul 2>&1")
    If Err.Number <> 0 Then
        IsAdmin = False
    Else
        IsAdmin = True
    End If
    On Error GoTo 0
End Function

' Function to relaunch script as admin
Sub RunAsAdmin()
    Dim shell, args, cmdLine
    Set shell = CreateObject("Shell.Application")
    args = """" & WScript.ScriptFullName & """"
    shell.ShellExecute "wscript.exe", args, "", "runas", 1
End Sub
