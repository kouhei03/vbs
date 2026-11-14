Option Explicit

Dim service, processes, process, processName, found, matchCount

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo TestProcessRunning.vbs <process-name>"
    WScript.Quit 2
End If

processName = Trim(WScript.Arguments(0))
If Len(processName) = 0 Then
    WScript.Echo "Process name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

found = False
matchCount = 0
For Each process In processes
    If StrComp(process.Name, processName, vbTextCompare) = 0 Then
        found = True
        matchCount = matchCount + 1
        WScript.Echo process.Name & " is running (PID " & process.ProcessId & ")."
    End If
Next

If Not found Then
    WScript.Echo processName & " is not running."
End If
WScript.Echo "Matching process count: " & matchCount
