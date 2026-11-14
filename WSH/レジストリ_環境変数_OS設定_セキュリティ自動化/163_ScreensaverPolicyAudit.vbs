' スクリーンセーバー関連設定を照会します。設定変更はしません。
Option Explicit
Dim shell, path, names, name, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Control Panel\Desktop\"
names = Array("ScreenSaveActive", "ScreenSaveTimeOut", "ScreenSaverIsSecure", "SCRNSAVE.EXE")
For Each name In names
    On Error Resume Next
    value = shell.RegRead(path & name)
    If Err.Number = 0 Then
        WScript.Echo name & "=" & value
    Else
        WScript.Echo name & "=(not set)"
    End If
    Err.Clear
    On Error GoTo 0
Next
