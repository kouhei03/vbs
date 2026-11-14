' Explorerの隠しファイル・拡張子表示設定を読み取ります。設定変更はしません。
Option Explicit
Dim shell, path, names, name, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
names = Array("Hidden", "HideFileExt", "ShowSuperHidden")
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
