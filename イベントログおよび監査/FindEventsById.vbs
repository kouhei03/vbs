Option Explicit

Dim logName, eventId, hours, timeFilter, service, events, eventItem

If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo FindEventsById.vbs <System|Application|Security> <event-id> <hours 1-168>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Log must be System, Application, or Security."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(1), 0, 65535, eventId) Then
    WScript.Echo "Event ID must be an integer from 0 to 65535."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(2), 1, 168, hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='" & logName & _
    "' AND EventCode=" & eventId & _
    " AND TimeGenerated >= '" & timeFilter.Value & "'")

WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Source" & vbTab & "Record"
For Each eventItem In events
    WScript.Echo FormatEventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.SourceName & vbTab & _
        eventItem.RecordNumber
Next

Function IsAllowedLog(name)
    IsAllowedLog = (StrComp(name, "System", vbTextCompare) = 0 Or _
        StrComp(name, "Application", vbTextCompare) = 0 Or _
        StrComp(name, "Security", vbTextCompare) = 0)
End Function

Function ParseInteger(text, minimum, maximum, value)
    Dim i, ch
    ParseInteger = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseInteger = (value >= minimum And value <= maximum)
End Function

Function FormatEventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    FormatEventTime = CStr(dateTime.GetVarDate(True))
End Function
