' ファイル概要: 自動ログオンの状態とユーザー名を確認する（パスワードは取得しない）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectAutoLogonState.vbs
' 実行結果: 自動ログオンの状態とユーザー名を確認する（パスワードは取得しない）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, autoLogon, userName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", autoLogon)
If result <> 0 Then autoLogon = "(not set)"
WScript.Echo "AutoAdminLogon: " & autoLogon
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName", userName)
If result <> 0 Then userName = "(not set)"
WScript.Echo "DefaultUserName: " & userName
WScript.Echo "Password value was not read or displayed."
