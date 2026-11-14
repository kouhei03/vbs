' ファイル概要: ローカルグループを一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListLocalGroups.vbs
' 実行結果: ローカルグループを一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, groups, group
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set groups = service.ExecQuery( _
    "SELECT Name, Description FROM Win32_Group WHERE LocalAccount=True")

WScript.Echo "Local group" & vbTab & "Description"
For Each group In groups
    WScript.Echo group.Name & vbTab & group.Description
Next
