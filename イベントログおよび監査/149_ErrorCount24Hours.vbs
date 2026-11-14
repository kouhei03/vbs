Option Explicit

Dim service, events, eventItem, filter, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent " & _
    "WHERE (Logfile='System' OR Logfile='Application') " & _
    "AND Type='Error' AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "System/Application error count in the last 24 hours: " & count
