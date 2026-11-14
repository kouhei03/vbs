' 概要: プロセスごとのCPUカウンターとprivate working setを一時点で表示します。
' 実行方法: cscript //nologo 059_ProcessResourceSnapshot.vbs
' 処理結果: CPUカウンターとprivate working setのスナップショット。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, IDProcess, PercentProcessorTime, WorkingSetPrivate FROM Win32_PerfFormattedData_PerfProc_Process")
WScript.Echo "Name" & vbTab & "PID" & vbTab & "CPU%" & vbTab & "PrivateWorkingSetBytes"
For Each item In records
    If item.IDProcess > 0 Then WScript.Echo item.Name & vbTab & item.IDProcess & vbTab & item.PercentProcessorTime & vbTab & item.WorkingSetPrivate
Next
