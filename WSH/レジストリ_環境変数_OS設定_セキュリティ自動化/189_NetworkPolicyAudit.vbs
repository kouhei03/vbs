' 概要: Windows TCP/IPの代表的なローカル設定値の存在を確認します。ロールバックはしません。
' 実行方法: cscript //nologo 189_NetworkPolicyAudit.vbs
' 処理結果: TCP/IP関連値の監査（ロールバックなし）。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\EnableICMPRedirect", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\DisableIPSourceRouting")
For Each path In paths
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
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
