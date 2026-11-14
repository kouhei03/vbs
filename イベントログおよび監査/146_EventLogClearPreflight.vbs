Option Explicit

Dim logName, service, logs, logItem, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 146_EventLogClearPreflight.vbs <System|Application|Security>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Unsupported event log."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set logs = service.ExecQuery( _
    "SELECT Name, FileSize, MaxFileSize, OverwritePolicy FROM Win32_NTEventLogFile " & _
    "WHERE Name='" & logName & "'")
found = False
For Each logItem In logs
    found = True
    WScript.Echo "Log: " & logItem.Name
    WScript.Echo "Current size: " & logItem.FileSize
    WScript.Echo "Maximum size: " & logItem.MaxFileSize
    WScript.Echo "Overwrite policy: " & logItem.OverwritePolicy
Next
If Not found Then WScript.Echo "Log metadata is unavailable."
WScript.Echo "No log entries were cleared. Preserve logs using approved retention procedures."

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function
