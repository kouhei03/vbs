Option Explicit

Dim hours, filter, service, events, eventItem
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 153_LogonLogoffTimeline.vbs <hours 1-168>"
    WScript.Quit 2
End If
If Not ParseHours(WScript.Arguments(0), hours) Then
    WScript.Echo "Hours must be an integer from 1 to 168."
    WScript.Quit 2
End If
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Security' AND (EventCode=4624 OR EventCode=4634 OR EventCode=4647) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
WScript.Echo "Time" & vbTab & "Event ID"
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode
Next
WScript.Echo "Account and network details are omitted."

Function ParseHours(text, value)
    Dim i, character
    ParseHours = False
    If Len(text) = 0 Or Len(text) > 3 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseHours = (value >= 1 And value <= 168)
End Function
