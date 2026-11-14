' ファイル概要: システム環境変数を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListSystemEnvironmentVariables.vbs
' 実行結果: システム環境変数を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, environment, variable
Set shell = CreateObject("WScript.Shell")
' [処理解説] SYSTEMスコープの環境変数を列挙します。値を変更しません。
Set environment = shell.Environment("SYSTEM")
For Each variable In environment
    WScript.Echo variable
Next
