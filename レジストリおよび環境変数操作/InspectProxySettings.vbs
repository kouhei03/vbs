' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: プロキシ設定値とレジストリ照会結果を保持します。
Dim registry, result, enabled, server, autoConfig, displayValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 現在のユーザーのプロキシ有効化、サーバー名、自動構成URLを読み取ります。
result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", enabled)
If result = 0 Then WScript.Echo "ProxyEnable=" & enabled
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", server)
If result = 0 Then
    displayValue = server
    If InStr(displayValue, "@") > 0 Then displayValue = "(value omitted because it may contain credentials)"
    WScript.Echo "ProxyServer=" & displayValue
End If
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "AutoConfigURL", autoConfig)
If result = 0 Then WScript.Echo "AutoConfigURL=" & autoConfig
