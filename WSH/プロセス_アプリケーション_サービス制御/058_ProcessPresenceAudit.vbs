' 不要アプリ候補の存在を確認するだけの監査例です。停止は行いません。
Option Explicit
Dim name, re, service, records, process
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 058_ProcessPresenceAudit.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ProcessId, CreationDate FROM Win32_Process WHERE Name='" & name & "'")
If records.Count = 0 Then WScript.Echo "No matching process: " & name
For Each process In records
    WScript.Echo "Candidate=" & name & " PID=" & process.ProcessId & " Created=" & process.CreationDate
Next
WScript.Echo "Audit only; no process was stopped."
