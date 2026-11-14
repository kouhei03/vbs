' サービスの現在のスタートアップ種類を照会します。変更は行いません。
Option Explicit
Dim name, re, service, records, item
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 055_ServiceStartupTypeAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, StartMode, State FROM Win32_Service WHERE Name='" & name & "'")
If records.Count = 0 Then WScript.Echo "Service not found: " & name: WScript.Quit 1
For Each item In records
    WScript.Echo item.Name & " StartMode=" & item.StartMode & " State=" & item.State
Next
