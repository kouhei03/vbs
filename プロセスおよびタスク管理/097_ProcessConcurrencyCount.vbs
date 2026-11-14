Option Explicit
Dim service, processes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT ProcessId FROM Win32_Process")
WScript.Echo "Current local process count=" & processes.Count
