' ファイル概要: Windowsサービスの名前と状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetWindowsServices.vbs
' 実行結果: Windowsサービスの名前と状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "State"
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
