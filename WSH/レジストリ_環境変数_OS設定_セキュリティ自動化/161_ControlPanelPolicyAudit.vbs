' 現在のユーザーのControl Panel制限ポリシーを照会します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoControlPanel"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "NoControlPanel=(not set)"
Else
    WScript.Echo "NoControlPanel=" & value & " (read-only)"
End If
On Error GoTo 0
