Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId
Next
