' 概要: Windows Updateの代表的なポリシー値の存在のみを確認します。設定は変更しません。
' 実行方法: cscript //nologo 159_WindowsUpdatePolicyAudit.vbs
' 処理結果: Windows Updateポリシー確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, values, names, name, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\"
names = Array("NoAutoUpdate", "AU\AUOptions", "SetDisableUXWUAccess")
For Each name In names
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
    value = shell.RegRead(path & name)
    If Err.Number = 0 Then
        WScript.Echo name & "=" & value
    Else
        WScript.Echo name & "=(not set or inaccessible)"
    End If
    Err.Clear
    On Error GoTo 0
Next
WScript.Echo "Read-only policy audit."
