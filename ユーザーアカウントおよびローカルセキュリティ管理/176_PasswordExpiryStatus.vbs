' ファイル概要: ローカルアカウントのパスワード有効期限フラグを確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 176_PasswordExpiryStatus.vbs
' 実行結果: ローカルアカウントのパスワード有効期限フラグを確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, accounts, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserAccount を照会し、Name, PasswordExpires, PasswordRequired を取得します。
Set accounts = service.ExecQuery( _
    "SELECT Name, PasswordExpires, PasswordRequired FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
WScript.Echo "Account" & vbTab & "PasswordExpires" & vbTab & "PasswordRequired"
For Each account In accounts
    WScript.Echo account.Name & vbTab & account.PasswordExpires & vbTab & _
        account.PasswordRequired
Next
WScript.Echo "Status only; password-expiry settings were not changed."
