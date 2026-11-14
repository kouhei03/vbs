' 概要: 現在のユーザーの通知許可状態を読み取ります。通知設定は変更しません。
' 実行方法: cscript //nologo 174_NotificationPolicyAudit.vbs
' 処理結果: 通知設定値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\PushNotifications\ToastEnabled"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "ToastEnabled=(not set or unavailable)"
Else
    WScript.Echo "ToastEnabled=" & value & " (read-only)"
End If
On Error GoTo 0
