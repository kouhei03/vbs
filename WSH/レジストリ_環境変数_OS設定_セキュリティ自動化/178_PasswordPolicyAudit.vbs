' ローカルアカウントのパスワード変更可否などを読み取ります。アカウント設定は変更しません。
Option Explicit
Dim service, accounts, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery("SELECT Name, LocalAccount, PasswordChangeable, PasswordExpires FROM Win32_UserAccount WHERE LocalAccount=True")
For Each account In accounts
    WScript.Echo account.Name & " PasswordChangeable=" & account.PasswordChangeable & _
        " PasswordExpires=" & account.PasswordExpires
Next
