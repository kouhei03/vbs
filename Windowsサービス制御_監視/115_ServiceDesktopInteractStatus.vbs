' ローカルサービスのDesktopInteract設定を照会します。
Option Explicit
Dim serviceName, validator
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 115_ServiceDesktopInteractStatus.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    WScript.Quit 2
End If
Dim service, services, item, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DesktopInteract FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        WScript.Echo "Name=" & item.Name
        WScript.Echo "DesktopInteract=" & item.DesktopInteract
        WScript.Echo "ExitCode=" & item.ExitCode
        found = True
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
