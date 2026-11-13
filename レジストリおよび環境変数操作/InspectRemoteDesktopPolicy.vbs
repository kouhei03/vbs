' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: RDP接続許可値を読み取るためのレジストリプロバイダーと結果を保持します。
Dim registry, result, denyConnections
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: Terminal ServicesのfDenyTSConnections値を読みます。RDP設定は変更しません。
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Control\Terminal Server", "fDenyTSConnections", denyConnections)
If result <> 0 Then
    WScript.Echo "RDP policy value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "fDenyTSConnections=" & denyConnections
