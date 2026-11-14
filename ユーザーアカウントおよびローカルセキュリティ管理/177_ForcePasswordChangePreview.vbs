' ファイル概要: 指定アカウントのパスワード状態を確認する（設定変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 177_ForcePasswordChangePreview.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定アカウントのパスワード状態を確認する（設定変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim userName, service, accounts, account, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 177_ForcePasswordChangePreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserAccount を照会し、Name, PasswordChangeable, PasswordExpires を取得します。
Set accounts = service.ExecQuery( _
    "SELECT Name, PasswordChangeable, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
found = False
For Each account In accounts
    If StrComp(account.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Account: " & account.Name
        WScript.Echo "Password changeable: " & account.PasswordChangeable
        WScript.Echo "Password expires: " & account.PasswordExpires
    End If
Next
If Not found Then WScript.Echo "Local account was not found."
WScript.Echo "Preview only; no next-logon flag was set."
