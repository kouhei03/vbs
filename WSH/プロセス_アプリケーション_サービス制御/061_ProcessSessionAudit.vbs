' 指定プロセスのSessionIdを照会します。別セッションでの起動は行いません。
Option Explicit
Dim name, re, service, records, process
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 061_ProcessSessionAudit.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ProcessId, SessionId FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo name & " PID=" & process.ProcessId & " SessionId=" & process.SessionId
Next
