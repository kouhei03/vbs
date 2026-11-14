' ファイル概要: ローカルコンピューター名を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetComputerName.vbs
' 実行結果: ローカルコンピューター名を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim network
Set network = CreateObject("WScript.Network")
WScript.Echo "Computer name: " & network.ComputerName
