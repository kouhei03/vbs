' プロセス優先度を照会します。優先度・CPU affinityの変更は行いません。
Option Explicit
Dim service, records, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, ProcessId, Priority FROM Win32_Process")
For Each process In records
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & "Priority=" & process.Priority
Next
WScript.Echo "Read-only audit; no scheduling settings changed."
