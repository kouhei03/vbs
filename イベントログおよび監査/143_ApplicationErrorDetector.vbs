Option Explicit

Dim service, events, eventItem, filter, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND (EventCode=1000 OR EventCode=1001) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")

count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & vbTab & _
        eventItem.SourceName & vbTab & eventItem.RecordNumber
    count = count + 1
Next
WScript.Echo "Application error/crash events in the last 24 hours: " & count
WScript.Echo "Event message text was not read."
