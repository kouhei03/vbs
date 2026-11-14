' ファイル概要: ローカルサービスの名前と状態を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetAllServices.vbs
' 実行結果: ローカルサービスの名前と状態を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State を取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "State"
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
