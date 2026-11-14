Option Explicit

Dim days, timeFilter, service, events, eventItem, count7031, count7034
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetServiceCrashCounts.vbs <days 1-30>"
    WScript.Quit 2
End If
If Not ParseDays(WScript.Arguments(0), days) Then
    WScript.Echo "Days must be a whole number from 1 to 30."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("d", -days, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='System' " & _
    "AND SourceName='Service Control Manager' " & _
    "AND (EventCode=7031 OR EventCode=7034) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

count7031 = 0
count7034 = 0
For Each eventItem In events
    If eventItem.EventCode = 7031 Then
        count7031 = count7031 + 1
    ElseIf eventItem.EventCode = 7034 Then
        count7034 = count7034 + 1
    End If
Next
WScript.Echo "Unexpected service termination events (last " & days & " days)"
WScript.Echo "Event 7031: " & count7031
WScript.Echo "Event 7034: " & count7034
WScript.Echo "Counts only; event message text is omitted."

Function ParseDays(text, value)
    Dim i, character
    ParseDays = False
    If Len(text) = 0 Or Len(text) > 2 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseDays = (value >= 1 And value <= 30)
End Function
