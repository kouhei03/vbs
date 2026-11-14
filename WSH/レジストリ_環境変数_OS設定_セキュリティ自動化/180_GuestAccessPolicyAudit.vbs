' 匿名／ゲスト共有に関する代表的なローカル設定値を照会します。変更はしません。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters\AllowInsecureGuestAuth", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\EveryoneIncludesAnonymous")
For Each path In paths
    On Error Resume Next
    value = shell.RegRead(path)
    If Err.Number = 0 Then
        WScript.Echo path & "=" & value
    Else
        WScript.Echo path & "=(not set or inaccessible)"
    End If
    Err.Clear
    On Error GoTo 0
Next
