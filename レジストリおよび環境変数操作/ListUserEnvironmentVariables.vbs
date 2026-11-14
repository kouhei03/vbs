' ファイル概要: 現在のユーザー環境変数を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListUserEnvironmentVariables.vbs
' 実行結果: 現在のユーザー環境変数を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, environment, variable
Set shell = CreateObject("WScript.Shell")
' [処理解説] 現在のユーザースコープの環境変数を列挙します。値を変更しません。
Set environment = shell.Environment("USER")
For Each variable In environment
    WScript.Echo variable
Next
