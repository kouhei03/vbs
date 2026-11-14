' プロセスごとのCPUカウンターとprivate working setを一時点で表示します。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, IDProcess, PercentProcessorTime, WorkingSetPrivate FROM Win32_PerfFormattedData_PerfProc_Process")
WScript.Echo "Name" & vbTab & "PID" & vbTab & "CPU%" & vbTab & "PrivateWorkingSetBytes"
For Each item In records
    If item.IDProcess > 0 Then WScript.Echo item.Name & vbTab & item.IDProcess & vbTab & item.PercentProcessorTime & vbTab & item.WorkingSetPrivate
Next
