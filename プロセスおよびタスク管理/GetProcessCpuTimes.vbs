Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, UserModeTime, KernelModeTime FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "User CPU time (100ns)" & vbTab & "Kernel CPU time (100ns)"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.UserModeTime & vbTab & process.KernelModeTime
Next
