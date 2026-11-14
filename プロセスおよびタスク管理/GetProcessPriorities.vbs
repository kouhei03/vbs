' ファイル概要: プロセスごとの優先度を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessPriorities.vbs
' 実行結果: プロセスごとの優先度を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, Priority FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Priority class"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.Priority
Next
