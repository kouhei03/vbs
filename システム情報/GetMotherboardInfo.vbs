' ファイル概要: マザーボードの製造元と製品名を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetMotherboardInfo.vbs
' 実行結果: マザーボードの製造元と製品名を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, boards, board
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_BaseBoard を照会し、Manufacturer, Product を取得します。
Set boards = service.ExecQuery( _
    "SELECT Manufacturer, Product FROM Win32_BaseBoard")

For Each board In boards
    WScript.Echo "Manufacturer: " & board.Manufacturer
    WScript.Echo "Model: " & board.Product
Next
