Option Explicit

Dim timeFilter, service, events, eventItem, count
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND (EventCode=1000 OR EventCode=1001) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Application crash/error events in the last 24 hours: " & count
