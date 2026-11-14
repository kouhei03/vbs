Option Explicit
Dim shell, environment
Set shell=CreateObject("WScript.Shell")
Set environment=shell.Environment("PROCESS")
environment("VBS_PROCESS_DEMO")="temporary-value"
WScript.Echo "Process-scope value=" & environment("VBS_PROCESS_DEMO")
WScript.Echo "This value is limited to this script process and its children."
