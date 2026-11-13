Option Explicit

Dim logName, hours, outputPath, timeFilter, service, events, eventItem
Dim fileSystem, outputFile

If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo ExportEventSummary.vbs <System|Application|Security> <hours 1-168> <new-csv-path>"
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

outputPath = WScript.Arguments(2)
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If fileSystem.FileExists(outputPath) Then
    WScript.Echo "Output file already exists; refusing to overwrite."
    WScript.Quit 2
End If

Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -hours, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='" & logName & _
    "' AND TimeGenerated >= '" & timeFilter.Value & "'")

On Error Resume Next
Set outputFile = fileSystem.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create CSV file: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

outputFile.WriteLine """Time"",""EventID"",""Source"",""RecordNumber"""
For Each eventItem In events
    outputFile.WriteLine Csv(EventTime(eventItem.TimeGenerated)) & "," & _
        Csv(eventItem.EventCode) & "," & Csv(eventItem.SourceName) & "," & _
        Csv(eventItem.RecordNumber)
Next
outputFile.Close
WScript.Echo "Event metadata exported; event message text was not included."

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

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function

Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
