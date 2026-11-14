Option Explicit

Dim service, services, item, nameUpper, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")
found = False

For Each item In services
    nameUpper = UCase(item.Name)
    If nameUpper = "W3SVC" Or nameUpper = "WAS" Or _
        nameUpper = "MSSQLSERVER" Or nameUpper = "SQLSERVERAGENT" Or _
        Left(nameUpper, 6) = "MSSQL$" Or Left(nameUpper, 9) = "SQLAGENT$" Then
        found = True
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
    End If
Next
If Not found Then WScript.Echo "No matching SQL Server or IIS services found."
