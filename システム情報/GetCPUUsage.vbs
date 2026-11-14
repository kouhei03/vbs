Option Explicit

Dim service, counters, counter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set counters = service.ExecQuery( _
    "SELECT Name, PercentProcessorTime FROM " & _
    "Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")

For Each counter In counters
    WScript.Echo "Current total CPU usage (%): " & counter.PercentProcessorTime
Next
