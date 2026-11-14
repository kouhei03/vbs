Option Explicit

Dim service, counters, counter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set counters = service.ExecQuery( _
    "SELECT Name, IDProcess, PercentProcessorTime " & _
    "FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess > 0")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "CPU counter (%)"
For Each counter In counters
    WScript.Echo counter.Name & vbTab & counter.IDProcess & vbTab & _
        counter.PercentProcessorTime
Next
