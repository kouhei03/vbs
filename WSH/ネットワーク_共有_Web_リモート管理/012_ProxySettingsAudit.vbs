' 概要: 現在のユーザープロキシ設定を読み取ります。レジストリ変更はしません。
' 実行方法: cscript //nologo 012_ProxySettingsAudit.vbs
' 処理結果: 現在のユーザープロキシ設定確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, enabled, server, autoUrl
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
enabled = shell.RegRead(path & "ProxyEnable")
If Err.Number <> 0 Then enabled = "(not set)": Err.Clear
server = shell.RegRead(path & "ProxyServer")
If Err.Number <> 0 Then server = "(not set)": Err.Clear
autoUrl = shell.RegRead(path & "AutoConfigURL")
If Err.Number <> 0 Then autoUrl = "(not set)": Err.Clear
On Error GoTo 0
WScript.Echo "ProxyEnable=" & enabled
WScript.Echo "ProxyServer=" & server
WScript.Echo "AutoConfigURL=" & autoUrl
