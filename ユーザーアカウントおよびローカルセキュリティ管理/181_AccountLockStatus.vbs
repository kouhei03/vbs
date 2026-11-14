' ファイル概要: アカウントのロックアウト状態を確認する（解除なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 181_AccountLockStatus.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: アカウントのロックアウト状態を確認する（解除なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim userName, service, accounts, account, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 181_AccountLockStatus.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery( _
    "SELECT Name, Lockout, Disabled FROM Win32_UserAccount WHERE LocalAccount=True")
found = False
For Each account In accounts
    If StrComp(account.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Account: " & account.Name
        WScript.Echo "Locked out: " & account.Lockout
        WScript.Echo "Disabled: " & account.Disabled
    End If
Next
If Not found Then WScript.Echo "Local account was not found."
WScript.Echo "Status only; no unlock operation was performed."
