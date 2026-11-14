' 現在のユーザーの通知許可状態を読み取ります。通知設定は変更しません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\PushNotifications\ToastEnabled"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "ToastEnabled=(not set or unavailable)"
Else
    WScript.Echo "ToastEnabled=" & value & " (read-only)"
End If
On Error GoTo 0
