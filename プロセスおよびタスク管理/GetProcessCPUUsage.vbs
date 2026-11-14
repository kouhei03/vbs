' ファイル概要: プロセスごとのCPU使用率を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessCPUUsage.vbs
' 実行結果: プロセスごとのCPU使用率を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, counters, counter
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PerfFormattedData_PerfProc_Process を照会し、Name, IDProcess, PercentProcessorTime を取得します。
Set counters = service.ExecQuery( _
    "SELECT Name, IDProcess, PercentProcessorTime " & _
    "FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess > 0")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "CPU counter (%)"
For Each counter In counters
    WScript.Echo counter.Name & vbTab & counter.IDProcess & vbTab & _
        counter.PercentProcessorTime
Next
