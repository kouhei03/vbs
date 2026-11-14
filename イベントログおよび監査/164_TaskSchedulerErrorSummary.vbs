Option Explicit

Dim filter, service, events, eventItem, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Microsoft-Windows-TaskScheduler/Operational' " & _
    "AND (EventCode=101 OR EventCode=103 OR EventCode=203) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode
    count = count + 1
Next
WScript.Echo "Task Scheduler error events in the last seven days: " & count
