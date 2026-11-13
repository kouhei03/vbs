Option Explicit

Dim service, processors, processor
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processors = service.ExecQuery( _
    "SELECT Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors " & _
    "FROM Win32_Processor")

For Each processor In processors
    WScript.Echo "Model: " & processor.Name
    WScript.Echo "Maximum clock speed (MHz): " & processor.MaxClockSpeed
    WScript.Echo "Cores: " & processor.NumberOfCores
    WScript.Echo "Logical processors: " & processor.NumberOfLogicalProcessors
    WScript.Echo ""
Next
