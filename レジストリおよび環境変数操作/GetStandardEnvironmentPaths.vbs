' ファイル概要: 標準のプロセス環境変数パスを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetStandardEnvironmentPaths.vbs
' 実行結果: 標準のプロセス環境変数パスを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shell, environment, variableName
Set shell = CreateObject("WScript.Shell")
' [処理解説] 環境変数は現在のプロセスのスコープを参照します。
Set environment = shell.Environment("PROCESS")

For Each variableName In Array("SystemRoot", "TEMP", "TMP")
    WScript.Echo variableName & "=" & environment(variableName)
Next
