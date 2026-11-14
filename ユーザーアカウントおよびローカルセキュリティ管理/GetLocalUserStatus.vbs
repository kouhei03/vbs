' ファイル概要: 指定したローカルユーザーの状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetLocalUserStatus.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定したローカルユーザーの状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, users, user, userName, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetLocalUserStatus.vbs <user-name>"
    WScript.Quit 2
End If
userName = Trim(WScript.Arguments(0))
If Len(userName) = 0 Then
    WScript.Echo "User name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
found = False
For Each user In users
    If StrComp(user.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "User: " & user.Name
        WScript.Echo "Disabled: " & user.Disabled
        WScript.Echo "Locked out: " & user.Lockout
        WScript.Echo "Password expires: " & user.PasswordExpires
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Local user not found: " & userName
    WScript.Quit 3
End If
