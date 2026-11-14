' 現在のユーザーのTask Manager制限ポリシーを照会します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "DisableTaskMgr=(not set)"
Else
    WScript.Echo "DisableTaskMgr=" & value & " (1=disabled, 0=enabled)"
End If
On Error GoTo 0
WScript.Echo "Registry was not modified."
