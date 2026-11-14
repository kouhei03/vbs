Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, ProcessId FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & _
    "State" & vbTab & "PID"
For Each item In services
    If item.ProcessId > 0 Then
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & _
            item.State & vbTab & item.ProcessId
    End If
Next
