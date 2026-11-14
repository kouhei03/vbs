Option Explicit

Dim service, logs, logItem, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize FROM Win32_NTEventLogFile " & _
    "WHERE Name='Security'")
found = False
For Each logItem In logs
    found = True
    WScript.Echo "Security log file: " & logItem.FileName
    WScript.Echo "Current bytes: " & logItem.FileSize
    WScript.Echo "Maximum bytes: " & logItem.MaxFileSize
Next
If Not found Then
    WScript.Echo "Security log metadata is unavailable; check permissions."
End If
WScript.Echo "Preflight only; no backup or export was created."
