' ファイル概要: 実行中のプロセス名とIDを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetRunningProcesses.vbs
' 実行結果: 実行中のプロセス名とIDを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId
Next
