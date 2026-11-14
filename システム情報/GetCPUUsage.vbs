' ファイル概要: CPU全体の使用率を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetCPUUsage.vbs
' 実行結果: CPU全体の使用率を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, counters, counter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set counters = service.ExecQuery( _
    "SELECT Name, PercentProcessorTime FROM " & _
    "Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")

For Each counter In counters
    WScript.Echo "Current total CPU usage (%): " & counter.PercentProcessorTime
Next
