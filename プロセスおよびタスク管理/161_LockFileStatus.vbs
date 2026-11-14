Option Explicit
Dim fso, shell, lockPath
Set fso=CreateObject("Scripting.FileSystemObject")
Set shell=CreateObject("WScript.Shell")
lockPath=shell.ExpandEnvironmentStrings("%TEMP%") & "\vbs-process-example.lock"
If fso.FileExists(lockPath) Then
    WScript.Echo "Lock file exists: " & lockPath
Else
    WScript.Echo "No lock file exists."
End If
