' 現在のユーザープロキシ設定を読み取ります。レジストリ変更はしません。
Option Explicit
Dim shell, path, enabled, server, autoUrl
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
On Error Resume Next
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
