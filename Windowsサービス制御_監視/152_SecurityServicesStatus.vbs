Option Explicit
Dim service, services, item, wantedNames, wantedName, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, ExitCode FROM Win32_Service")
wantedNames = Array("WinDefend", "MpsSvc", "wscsvc", "WdNisSvc")
For Each wantedName In wantedNames
    found = False
    For Each item In services
        If StrComp(item.Name, wantedName, vbTextCompare) = 0 Then
            WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode & vbTab & item.ExitCode
            found = True
            Exit For
        End If
    Next
    If Not found Then WScript.Echo wantedName & vbTab & "Not installed or not visible"
Next
