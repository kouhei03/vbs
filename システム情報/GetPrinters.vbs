' ファイル概要: プリンターの名前と接続状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetPrinters.vbs
' 実行結果: プリンターの名前と接続状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, printers, printer
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Printer を照会し、Name, Default, Network, Local を取得します。
Set printers = service.ExecQuery( _
    "SELECT Name, Default, Network, Local FROM Win32_Printer")

For Each printer In printers
    WScript.Echo "Printer: " & printer.Name
    WScript.Echo "Default: " & printer.Default
    WScript.Echo "Network printer: " & printer.Network
    WScript.Echo "Local printer: " & printer.Local
    WScript.Echo ""
Next
