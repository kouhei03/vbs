' Windows TCP/IPの代表的なローカル設定値の存在を確認します。ロールバックはしません。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\EnableICMPRedirect", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\DisableIPSourceRouting")
For Each path In paths
    On Error Resume Next
    value = shell.RegRead(path)
    If Err.Number = 0 Then
        WScript.Echo path & "=" & value
    Else
        WScript.Echo path & "=(default or inaccessible)"
    End If
    Err.Clear
    On Error GoTo 0
Next
WScript.Echo "No registry values were reset."
