' ファイル概要: プロセスごとのプライベートメモリ使用量を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessPrivateMemory.vbs
' 実行結果: プロセスごとのプライベートメモリ使用量を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process, bytes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, PrivatePageCount FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Private memory (bytes)"
For Each process In processes
    bytes = process.PrivatePageCount
    If IsNull(bytes) Then bytes = "Unavailable"
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & bytes
Next
