' ファイル概要: ローカル共有の名前・パス・種類を表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 191_NetworkShareInventory.vbs
' 実行結果: ローカル共有の名前・パス・種類を表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, shares, share
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Share を照会し、Name, Path, Type, Description を取得します。
Set shares = service.ExecQuery( _
    "SELECT Name, Path, Type, Description FROM Win32_Share")
WScript.Echo "Name" & vbTab & "Path" & vbTab & "Type" & vbTab & "Description"
For Each share In shares
    WScript.Echo share.Name & vbTab & share.Path & vbTab & _
        share.Type & vbTab & share.Description
Next
