Option Explicit

Dim hours, timeFilter, service, events, eventItem

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetBootShutdownHistory.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be a whole number from 1 to 168."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='System' " & _
    "AND (EventCode=41 OR EventCode=12 OR EventCode=13 OR " & _
    "EventCode=6005 OR EventCode=6006 OR EventCode=6008) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Source" & vbTab & "Record"
For Each eventItem In events
    WScript.Echo EventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.SourceName & vbTab & _
        eventItem.RecordNumber
Next

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

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function
