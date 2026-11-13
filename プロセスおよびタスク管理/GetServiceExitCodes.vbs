Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "State" & vbTab & _
    "Exit code" & vbTab & "Service-specific exit code"
For Each item In services
    WScript.Echo item.Name & vbTab & item.State & vbTab & _
        item.ExitCode & vbTab & item.ServiceSpecificExitCode
Next
