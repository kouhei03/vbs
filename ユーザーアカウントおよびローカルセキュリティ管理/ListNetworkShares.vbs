' ファイル概要: ローカルコンピューターの共有リソースを一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListNetworkShares.vbs
' 実行結果: ローカルコンピューターの共有リソースを一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, shares, share
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set shares = service.ExecQuery( _
    "SELECT Name, Path, Type, Description FROM Win32_Share")

WScript.Echo "Share name" & vbTab & "Path" & vbTab & "Type" & vbTab & "Description"
For Each share In shares
    WScript.Echo share.Name & vbTab & share.Path & vbTab & _
        share.Type & vbTab & share.Description
Next
