' 概要: RDP関連のローカル設定値を読み取ります。レジストリ・ファイアウォールは変更しません。
' 実行方法: cscript //nologo 029_RdpConfigurationAudit.vbs
' 処理結果: RDP関連レジストリとFirewall状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim shell, value, keyPath, profiles, profile
Set shell = CreateObject("WScript.Shell")
keyPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\fDenyTSConnections"
On Error Resume Next
value = shell.RegRead(keyPath)
If Err.Number <> 0 Then
    WScript.Echo "RDP registry value unavailable: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "fDenyTSConnections=" & value & " (1=deny, 0=allow)"
' 主処理: WMIの照会結果を処理します。
Set profiles = GetObject("winmgmts:\\.\root\StandardCimv2").ExecQuery("SELECT Name, Enabled FROM MSFT_NetFirewallProfile")
For Each profile In profiles
    WScript.Echo "FirewallProfile=" & profile.Name & " Enabled=" & profile.Enabled
Next
