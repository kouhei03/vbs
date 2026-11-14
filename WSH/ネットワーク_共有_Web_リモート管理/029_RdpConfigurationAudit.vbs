' RDP関連のローカル設定値を読み取ります。レジストリ・ファイアウォールは変更しません。
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
Set profiles = GetObject("winmgmts:\\.\root\StandardCimv2").ExecQuery("SELECT Name, Enabled FROM MSFT_NetFirewallProfile")
For Each profile In profiles
    WScript.Echo "FirewallProfile=" & profile.Name & " Enabled=" & profile.Enabled
Next
