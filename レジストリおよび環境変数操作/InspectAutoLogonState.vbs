' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: レジストリプロバイダー、認証設定、戻り値を保持します。
Dim registry, result, autoLogon, userName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 自動ログオンの有効フラグとユーザー名だけを読み、パスワード値にはアクセスしません。
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", autoLogon)
If result <> 0 Then autoLogon = "(not set)"
WScript.Echo "AutoAdminLogon: " & autoLogon
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName", userName)
If result <> 0 Then userName = "(not set)"
WScript.Echo "DefaultUserName: " & userName
WScript.Echo "Password value was not read or displayed."
