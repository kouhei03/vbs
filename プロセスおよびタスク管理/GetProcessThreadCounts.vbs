' ファイル概要: プロセスごとのスレッド数を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessThreadCounts.vbs
' 実行結果: プロセスごとのスレッド数を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, ThreadCount を取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ThreadCount FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Threads"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.ThreadCount
Next
