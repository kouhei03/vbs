' ファイル概要: 共有名と対応するパスを確認する（削除なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 193_DeleteSharePreflight.vbs <share-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 共有名と対応するパスを確認する（削除なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shareName, service, shares, share, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 193_DeleteSharePreflight.vbs <share-name>"
    WScript.Quit 2
End If
shareName = WScript.Arguments(0)
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set shares = service.ExecQuery("SELECT Name, Path FROM Win32_Share")
found = False
For Each share In shares
    If StrComp(share.Name, shareName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Share: " & share.Name
        WScript.Echo "Backing path: " & share.Path
    End If
Next
If Not found Then WScript.Echo "Share was not found."
WScript.Echo "Preflight only; the share and its files were not deleted."
