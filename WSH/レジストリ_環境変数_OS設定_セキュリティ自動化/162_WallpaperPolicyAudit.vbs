' ユーザー壁紙と壁紙ポリシーの設定値を照会します。壁紙は変更しません。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_CURRENT_USER\Control Panel\Desktop\WallPaper", _
    "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\Wallpaper")
For Each path In paths
    On Error Resume Next
    value = shell.RegRead(path)
    If Err.Number = 0 Then
        WScript.Echo path & "=" & value
    Else
        WScript.Echo path & "=(not set)"
    End If
    Err.Clear
    On Error GoTo 0
Next
