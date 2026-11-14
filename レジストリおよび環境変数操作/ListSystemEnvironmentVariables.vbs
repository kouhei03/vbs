Option Explicit
Dim shell, environment, variable
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("SYSTEM")
For Each variable In environment
    WScript.Echo variable
Next
