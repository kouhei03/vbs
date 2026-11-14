' 概要: 匿名／ゲスト共有に関する代表的なローカル設定値を照会します。変更はしません。
' 実行方法: cscript //nologo 180_GuestAccessPolicyAudit.vbs
' 処理結果: ゲスト／匿名アクセス関連値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters\AllowInsecureGuestAuth", _
    "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa\EveryoneIncludesAnonymous")
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
