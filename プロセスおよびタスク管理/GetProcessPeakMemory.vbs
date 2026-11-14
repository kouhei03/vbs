' ファイル概要: プロセスごとのピークメモリ使用量を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessPeakMemory.vbs
' 実行結果: プロセスごとのピークメモリ使用量を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, PeakWorkingSetSize, PeakPageFileUsage を取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, PeakWorkingSetSize, PeakPageFileUsage " & _
    "FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Peak working set (bytes)" & vbTab & "Peak page file (KB)"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.PeakWorkingSetSize & vbTab & process.PeakPageFileUsage
Next
