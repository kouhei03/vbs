' 指定プロセスの親子関係を表示します。プロセス終了は行いません。
Option Explicit
Dim name, re, service, records, process, names, parents, key, pid
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 067_ProcessTreeAudit.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, ProcessId, ParentProcessId FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo "Name=" & process.Name & " PID=" & process.ProcessId & " ParentPID=" & process.ParentProcessId
Next
WScript.Echo "Audit only; process tree was not modified."
