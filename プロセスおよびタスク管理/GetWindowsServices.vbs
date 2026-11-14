Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "State"
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
