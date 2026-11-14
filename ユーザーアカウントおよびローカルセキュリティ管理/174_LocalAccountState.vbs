' ファイル概要: ローカルアカウントの無効化・ロック・パスワード状態を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 174_LocalAccountState.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ローカルアカウントの無効化・ロック・パスワード状態を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim userName, service, accounts, account, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 174_LocalAccountState.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires, PasswordChangeable " & _
    "FROM Win32_UserAccount WHERE LocalAccount=True")
found = False
For Each account In accounts
    If StrComp(account.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Account: " & account.Name
        WScript.Echo "Disabled: " & account.Disabled
        WScript.Echo "Locked out: " & account.Lockout
        WScript.Echo "Password expires: " & account.PasswordExpires
        WScript.Echo "Password changeable: " & account.PasswordChangeable
    End If
Next
If Not found Then WScript.Echo "Local account was not found."
