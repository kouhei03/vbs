Option Explicit

Dim service, services, item, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, CheckPoint, WaitHint " & _
    "FROM Win32_Service")

found = False
For Each item In services
    If StrComp(item.State, "Start Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Stop Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Continue Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Pause Pending", vbTextCompare) = 0 Then
        found = True
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & _
            item.State & vbTab & "Checkpoint=" & item.CheckPoint & _
            vbTab & "WaitHint(ms)=" & item.WaitHint
    End If
Next

If Not found Then WScript.Echo "No services are in a pending state."
