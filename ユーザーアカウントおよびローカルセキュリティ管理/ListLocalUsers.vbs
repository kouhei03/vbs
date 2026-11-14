' ファイル概要: ローカルユーザーとその状態を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListLocalUsers.vbs
' 実行結果: ローカルユーザーとその状態を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, users, user
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserAccount を照会し、Name, Disabled, Lockout, PasswordExpires を取得します。
Set users = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")

WScript.Echo "Local user" & vbTab & "Disabled" & vbTab & _
    "Locked out" & vbTab & "Password expires"
For Each user In users
    WScript.Echo user.Name & vbTab & user.Disabled & vbTab & _
        user.Lockout & vbTab & user.PasswordExpires
Next
