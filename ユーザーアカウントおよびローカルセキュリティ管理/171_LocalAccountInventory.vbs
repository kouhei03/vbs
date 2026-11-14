' ファイル概要: ローカルアカウントの名前・SID・状態を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 171_LocalAccountInventory.vbs
' 実行結果: ローカルアカウントの名前・SID・状態を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, users, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Name, SID, Disabled, Lockout FROM Win32_UserAccount WHERE LocalAccount=True")
WScript.Echo "Name" & vbTab & "SID" & vbTab & "Disabled" & vbTab & "Locked"
For Each account In users
    WScript.Echo account.Name & vbTab & account.SID & vbTab & _
        account.Disabled & vbTab & account.Lockout
Next
