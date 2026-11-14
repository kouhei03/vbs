Option Explicit

Dim timeFilter, service, events, eventItem
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='System' " & _
    "AND (EventCode=27 OR EventCode=4201) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

WScript.Echo "Potential network adapter state changes (last 24 hours)"
WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Source" & vbTab & "Record"
For Each eventItem In events
    WScript.Echo EventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.SourceName & vbTab & _
        eventItem.RecordNumber
Next

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function
