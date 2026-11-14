' 視覚効果プリセットの値だけを読み取ります。デザイン設定は変更しません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects\VisualFXSetting"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "VisualFXSetting=(Windows default or not set)"
Else
    WScript.Echo "VisualFXSetting=" & value & " (read-only)"
End If
On Error GoTo 0
