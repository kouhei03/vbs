Option Explicit

Dim service, services, item, serviceName
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo TestServiceExists.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
If Len(serviceName) = 0 Then
    WScript.Echo "Service name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        WScript.Echo "Service exists: " & item.Name
        WScript.Quit 0
    End If
Next
WScript.Echo "Service not found: " & serviceName
WScript.Quit 3
