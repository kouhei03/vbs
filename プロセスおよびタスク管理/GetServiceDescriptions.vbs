Option Explicit

Dim service, services, item, description
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, Description FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "Description"
For Each item In services
    description = item.Description
    If IsNull(description) Then
        description = "(no description)"
    ElseIf Len(description) = 0 Then
        description = "(no description)"
    End If
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & description
Next
