' ファイル概要: 無効サービスの復旧検討用一覧。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 122_DisabledServicesRestorePreview.vbs
' 実行結果: 無効サービスの復旧検討用一覧処理の結果やエラーをコンソールに表示します。

' 無効状態のサービスを列挙し、復旧検討用の情報を表示します。設定は変更しません。
Option Explicit
Dim service, services, item, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, StartMode を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode FROM Win32_Service WHERE StartMode='Disabled'")
count = 0
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode
    count = count + 1
Next
WScript.Echo "Disabled services found=" & count & "; no startup settings were changed."
