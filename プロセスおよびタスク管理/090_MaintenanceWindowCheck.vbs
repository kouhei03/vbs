Option Explicit
Dim startMinute, endMinute, nowMinute, insideWindow
If WScript.Arguments.Count <> 2 Then WScript.Echo "Usage: cscript //nologo 090_MaintenanceWindowCheck.vbs <start-minute-0-1439> <end-minute-0-1439>": WScript.Quit 2
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Times must be minute numbers.": WScript.Quit 2
startMinute = CLng(WScript.Arguments(0))
endMinute = CLng(WScript.Arguments(1))
If startMinute < 0 Or startMinute > 1439 Or endMinute < 0 Or endMinute > 1439 Then WScript.Echo "Minutes must be between 0 and 1439.": WScript.Quit 2
nowMinute = Hour(Now) * 60 + Minute(Now)
If startMinute <= endMinute Then
    insideWindow = (nowMinute >= startMinute And nowMinute <= endMinute)
Else
    insideWindow = (nowMinute >= startMinute Or nowMinute <= endMinute)
End If
WScript.Echo "Current minute=" & nowMinute & " Maintenance window active=" & insideWindow
