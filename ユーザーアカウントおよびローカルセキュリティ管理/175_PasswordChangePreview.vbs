' ファイル概要: パスワード変更対象のアカウント名を表示する（変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 175_PasswordChangePreview.vbs <user-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: パスワード変更対象のアカウント名を表示する（変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim userName
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 175_PasswordChangePreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
If Len(userName) = 0 Or Len(userName) > 256 Then
    WScript.Echo "User name length is invalid."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no password was changed."
WScript.Echo "Target local account: " & userName
WScript.Echo "Do not pass passwords on a command line."
