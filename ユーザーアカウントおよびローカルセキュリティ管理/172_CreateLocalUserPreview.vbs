' ファイル概要: 作成予定のユーザー名を検証する（作成なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 172_CreateLocalUserPreview.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 作成予定のユーザー名を検証する（作成なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim userName, expression
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 172_CreateLocalUserPreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9._-]{0,19}$"
If Not expression.Test(userName) Then
    WScript.Echo "User name must be 1-20 safe ASCII characters."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no account was created."
WScript.Echo "Proposed local account name: " & userName
WScript.Echo "No password was accepted or displayed."
