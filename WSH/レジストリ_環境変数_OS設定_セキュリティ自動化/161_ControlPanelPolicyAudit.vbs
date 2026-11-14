' 概要: 現在のユーザーのControl Panel制限ポリシーを照会します。
' 実行方法: cscript //nologo 161_ControlPanelPolicyAudit.vbs
' 処理結果: Control Panel制限ポリシー確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoControlPanel"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "NoControlPanel=(not set)"
Else
    WScript.Echo "NoControlPanel=" & value & " (read-only)"
End If
On Error GoTo 0
