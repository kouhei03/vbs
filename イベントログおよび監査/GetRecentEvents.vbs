Option Explicit

Dim logName, hours, timeFilter, service, events, eventItem

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo GetRecentEvents.vbs <System|Application|Security> <hours 1-168>"
    WScript.Quit 2
End If
logName = WScript.Arguments(0)
If Not IsAllowedLog(logName) Then
    WScript.Echo "Log must be System, Application, or Security."
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(1), hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='" & logName & _
    "' AND TimeGenerated >= '" & timeFilter.Value & "'")

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

Function ParseHours(text, value)
    Dim i, ch
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If InStr(1, "0123456789", ch, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 168)
End Function

Function FormatEventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    FormatEventTime = CStr(dateTime.GetVarDate(True))
End Function
