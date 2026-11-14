' ファイル概要: プロセスごとのセッションIDを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessSessionIds.vbs
' 実行結果: プロセスごとのセッションIDを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, SessionId FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Session ID"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.SessionId
Next
