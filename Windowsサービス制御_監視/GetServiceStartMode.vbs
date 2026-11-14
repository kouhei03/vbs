Option Explicit

Dim service, services, item, serviceName, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetServiceStartMode.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
If Len(serviceName) = 0 Then
    WScript.Echo "Service name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, StartMode FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo item.DisplayName & " (" & item.Name & _
            ") startup mode: " & item.StartMode
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
