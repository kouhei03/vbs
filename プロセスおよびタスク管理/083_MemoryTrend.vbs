Option Explicit
Dim service, records, process, processName, validator, i, found
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 083_MemoryTrend.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 5
    Set records = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process WHERE Name = '" & processName & "'")
    found = False
    For Each process In records
        WScript.Echo "Sample=" & i & " PID=" & process.ProcessId & " WorkingSetBytes=" & process.WorkingSetSize
        found = True
    Next
    If Not found Then WScript.Echo "Process not running at sample " & i
    If i < 5 Then WScript.Sleep 2000
Next
WScript.Echo "Samples are diagnostic only and do not assert a memory leak."
