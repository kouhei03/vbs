' ファイル概要: プロセススコープ環境変数のデモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 156_ProcessEnvironmentDemo.vbs
' 実行結果: プロセススコープ環境変数のデモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, environment
Set shell=CreateObject("WScript.Shell")
Set environment=shell.Environment("PROCESS")
environment("VBS_PROCESS_DEMO")="temporary-value"
WScript.Echo "Process-scope value=" & environment("VBS_PROCESS_DEMO")
WScript.Echo "This value is limited to this script process and its children."
