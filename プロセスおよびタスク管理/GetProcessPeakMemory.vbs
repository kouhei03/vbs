Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, PeakWorkingSetSize, PeakPageFileUsage " & _
    "FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Peak working set (bytes)" & vbTab & "Peak page file (KB)"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.PeakWorkingSetSize & vbTab & process.PeakPageFileUsage
Next
