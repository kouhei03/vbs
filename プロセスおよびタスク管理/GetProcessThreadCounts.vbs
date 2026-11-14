Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ThreadCount FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Threads"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.ThreadCount
Next
