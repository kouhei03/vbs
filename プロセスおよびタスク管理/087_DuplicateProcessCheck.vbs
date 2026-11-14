Option Explicit
Dim service, records, processName, validator
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 087_DuplicateProcessCheck.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
WScript.Echo "Running instances: " & records.Count
If records.Count > 1 Then WScript.Echo "Duplicate instances detected; no process was blocked or terminated."
