Option Explicit

Dim service, logs, logItem
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize, " & _
    "OverwriteOutdated, OverwritePolicy FROM Win32_NTEventLogFile")

WScript.Echo "Log" & vbTab & "Current size (bytes)" & vbTab & _
    "Maximum size (bytes)" & vbTab & "Overwrite policy"
For Each logItem In logs
    WScript.Echo logItem.Name & vbTab & logItem.FileSize & vbTab & _
        logItem.MaxFileSize & vbTab & logItem.OverwritePolicy
Next
