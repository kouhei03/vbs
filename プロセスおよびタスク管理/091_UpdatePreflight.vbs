Option Explicit
Dim service, processName, validator, records, i
If WScript.Arguments.Count < 1 Then WScript.Echo "Usage: cscript //nologo 091_UpdatePreflight.vbs <process.exe> [process.exe ...]": WScript.Quit 2
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 0 To WScript.Arguments.Count - 1
    processName = WScript.Arguments(i)
    If Not validator.Test(processName) Then WScript.Echo "Invalid name: " & processName: WScript.Quit 2
    Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    WScript.Echo processName & " running instances=" & records.Count
Next
WScript.Echo "Preflight only; no application was closed."
