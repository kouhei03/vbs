' 概要: ローカルアカウントのパスワード変更可否などを読み取ります。アカウント設定は変更しません。
' 実行方法: cscript //nologo 178_PasswordPolicyAudit.vbs
' 処理結果: ローカルアカウントのパスワード変更ポリシー確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, accounts, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set accounts = service.ExecQuery("SELECT Name, LocalAccount, PasswordChangeable, PasswordExpires FROM Win32_UserAccount WHERE LocalAccount=True")
For Each account In accounts
    WScript.Echo account.Name & " PasswordChangeable=" & account.PasswordChangeable & _
        " PasswordExpires=" & account.PasswordExpires
Next
