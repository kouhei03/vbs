' 指定した実行ファイル名のプロセスがローカルで動作中か確認します。
Option Explicit
Dim name, re, service, records, process, count
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 051_CheckProcessRunning.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='" & name & "'")
count = 0
For Each process In records
    WScript.Echo name & " PID=" & process.ProcessId
    count = count + 1
Next
If count = 0 Then
    WScript.Echo "Not running: " & name
    WScript.Quit 1
End If
