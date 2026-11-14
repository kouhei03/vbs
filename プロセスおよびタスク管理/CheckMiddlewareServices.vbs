Option Explicit

Dim service, services, item
Dim matched, upperName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")

matched = False
For Each item In services
    upperName = UCase(item.Name)
    If upperName = "W3SVC" Or upperName = "WAS" Or _
        upperName = "MSSQLSERVER" Or upperName = "SQLSERVERAGENT" Or _
        Left(upperName, 6) = "MSSQL$" Or Left(upperName, 9) = "SQLAGENT$" Then
        matched = True
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
    End If
Next

If Not matched Then
    WScript.Echo "No matching SQL Server or IIS services were found."
End If
