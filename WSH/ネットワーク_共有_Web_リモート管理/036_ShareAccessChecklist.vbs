' 指定した単一UNC共有へのアクセスを確認し、読み取り専用の診断項目を表示します。
Option Explicit
Dim path, re, fso, shell, runner
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 036_ShareAccessChecklist.vbs <UNC-share-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^\\\\[A-Za-z0-9.-]+\\[A-Za-z0-9_$.-]+(?:\\[A-Za-z0-9 _.$-]+)*$"
If Not re.Test(path) Then
    WScript.Echo "Provide one simple UNC share path."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
WScript.Echo "SharePath=" & path
If fso.FolderExists(path) Then
    WScript.Echo "Share access: available"
Else
    WScript.Echo "Share access: unavailable or access denied"
End If
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("net.exe use")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
WScript.Echo "Also check name resolution, VPN state, server availability, and assigned permissions."
