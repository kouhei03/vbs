' 指定Windowsサービスの状態を照会します。開始・停止・再起動は行いません。
Option Explicit
Dim name, re, service, records, item
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 054_ServiceStatusAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, State, StartMode, ProcessId FROM Win32_Service WHERE Name='" & name & "'")
If records.Count = 0 Then
    WScript.Echo "Service not found: " & name
    WScript.Quit 1
End If
For Each item In records
    WScript.Echo "Name=" & item.Name & " State=" & item.State & " StartMode=" & item.StartMode & " PID=" & item.ProcessId
Next
