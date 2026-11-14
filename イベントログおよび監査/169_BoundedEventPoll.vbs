Option Explicit

Dim eventId, minutes, intervalSeconds, service, filter, events
Dim eventItem, elapsed, count, previousCount
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 169_BoundedEventPoll.vbs <event-id> <minutes 1-5> <interval-seconds 5-30>"
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(0), 0, 65535, eventId) Then
    WScript.Echo "Event ID must be an integer from 0 to 65535."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(1), 1, 5, minutes) Then
    WScript.Echo "Minutes must be from 1 to 5."
    WScript.Quit 2
End If
If Not ParseInteger(WScript.Arguments(2), 5, 30, intervalSeconds) Then
    WScript.Echo "Interval must be from 5 to 30 seconds."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
previousCount = -1
For elapsed = 0 To minutes * 60 Step intervalSeconds
    Set filter = CreateObject("WbemScripting.SWbemDateTime")
    filter.SetVarDate DateAdd("s", -elapsed, Now), True
    Set events = service.ExecQuery( _
        "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='System' " & _
        "AND EventCode=" & eventId & " AND TimeGenerated >= '" & filter.Value & "'")
    count = 0
    For Each eventItem In events
        count = count + 1
    Next
    If count <> previousCount Then
        WScript.Echo "Elapsed seconds=" & elapsed & "; matching events in window=" & count
        previousCount = count
    End If
    If elapsed < minutes * 60 Then WScript.Sleep intervalSeconds * 1000
Next
WScript.Echo "Bounded observation ended."

Function ParseInteger(text, minimum, maximum, value)
    Dim i, character
    ParseInteger = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseInteger = (value >= minimum And value <= maximum)
End Function
