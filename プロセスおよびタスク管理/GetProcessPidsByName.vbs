Option Explicit

Dim service, processes, process, targetName, found

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetProcessPidsByName.vbs <process-name>"
    WScript.Quit 2
End If
targetName = Trim(WScript.Arguments(0))
If Len(targetName) = 0 Then
    WScript.Echo "Process name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

found = False
For Each process In processes
    If StrComp(process.Name, targetName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo process.Name & vbTab & process.ProcessId
    End If
Next
If Not found Then WScript.Echo "No matching process found."
