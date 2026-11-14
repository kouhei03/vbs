' ファイル概要: 同時プロセス数の取得。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 097_ProcessConcurrencyCount.vbs
' 実行結果: 同時プロセス数の取得処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT ProcessId FROM Win32_Process")
WScript.Echo "Current local process count=" & processes.Count
