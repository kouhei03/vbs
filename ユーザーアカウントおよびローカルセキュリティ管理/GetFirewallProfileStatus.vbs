' ファイル概要: ファイアウォールプロファイルの状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetFirewallProfileStatus.vbs
' 実行結果: ファイアウォールプロファイルの状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\standardcimv2")
' [処理解説] WMIの MSFT_NetFirewallProfile を照会し、Name, Enabled を取得します。
Set profiles = service.ExecQuery( _
    "SELECT Name, Enabled FROM MSFT_NetFirewallProfile")

WScript.Echo "Firewall profile" & vbTab & "Enabled"
For Each profile In profiles
    WScript.Echo profile.Name & vbTab & profile.Enabled
Next
