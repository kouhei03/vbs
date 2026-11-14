' 概要: AutoLogon関連設定を監査します。DefaultPasswordなどの秘密値は読み取りません。
' 実行方法: cscript //nologo 157_AutoLogonAudit.vbs
' 処理結果: 自動ログオン状態確認（パスワード値は非読取）。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, root, enabled, userName, passwordPresent
Set shell = CreateObject("WScript.Shell")
root = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
enabled = shell.RegRead(root & "AutoAdminLogon")
If Err.Number <> 0 Then enabled = "(not set)": Err.Clear
userName = shell.RegRead(root & "DefaultUserName")
If Err.Number <> 0 Then userName = "(not set)": Err.Clear
shell.RegRead root & "DefaultPassword"
passwordPresent = (Err.Number = 0)
Err.Clear
On Error GoTo 0
WScript.Echo "AutoAdminLogon=" & enabled
WScript.Echo "DefaultUserName=" & userName
WScript.Echo "DefaultPasswordPresent=" & passwordPresent & " (value not read)"
