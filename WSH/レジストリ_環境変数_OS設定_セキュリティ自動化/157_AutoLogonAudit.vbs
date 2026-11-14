' AutoLogon関連設定を監査します。DefaultPasswordなどの秘密値は読み取りません。
Option Explicit
Dim shell, root, enabled, userName, passwordPresent
Set shell = CreateObject("WScript.Shell")
root = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
On Error Resume Next
enabled = shell.RegRead(root & "AutoAdminLogon")
If Err.Number <> 0 Then enabled = "(not set)": Err.Clear
userName = shell.RegRead(root & "DefaultUserName")
If Err.Number <> 0 Then userName = "(not set)": Err.Clear
shell.RegRead root & "DefaultPassword"
passwordPresent = (Err.Number = 0)
Err.Clear
On Error GoTo 0
WScript.Echo "AutoAdminLogon=" & enabled
WScript.Echo "DefaultUserName=" & userName
WScript.Echo "DefaultPasswordPresent=" & passwordPresent & " (value not read)"
