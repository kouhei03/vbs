' 高速スタートアップ関連のレジストリ値を照会します。電源設定は変更しません。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\HibernateEnabled")
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
