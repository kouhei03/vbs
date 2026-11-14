Option Explicit

Dim service, logs, logItem
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize, OverwritePolicy, OverwriteOutdated " & _
    "FROM Win32_NTEventLogFile")
WScript.Echo "Log" & vbTab & "CurrentBytes" & vbTab & "MaximumBytes" & vbTab & "OverwritePolicy"
For Each logItem In logs
    WScript.Echo logItem.Name & vbTab & logItem.FileSize & vbTab & _
        logItem.MaxFileSize & vbTab & logItem.OverwritePolicy
Next
WScript.Echo "Read-only query; log settings were not changed."
