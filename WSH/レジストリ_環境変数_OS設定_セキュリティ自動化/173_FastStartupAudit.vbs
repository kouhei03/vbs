' 概要: 高速スタートアップ関連のレジストリ値を照会します。電源設定は変更しません。
' 実行方法: cscript //nologo 173_FastStartupAudit.vbs
' 処理結果: 高速スタートアップ関連値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power\HiberbootEnabled", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Power\HibernateEnabled")
For Each path In paths
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
    value = shell.RegRead(path)
    If Err.Number = 0 Then
        WScript.Echo path & "=" & value
    Else
        WScript.Echo path & "=(not set or inaccessible)"
    End If
    Err.Clear
    On Error GoTo 0
Next
