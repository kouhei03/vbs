' ファイル概要: 既存ウィンドウのアクティブ化。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 088_ActivateExistingWindow.vbs <window-title>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 既存ウィンドウのアクティブ化処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, title, activated
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 088_ActivateExistingWindow.vbs <window-title>": WScript.Quit 2
title = WScript.Arguments(0)
If Len(title) = 0 Then WScript.Echo "Window title must not be empty.": WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
activated = shell.AppActivate(title)
If activated Then
    WScript.Echo "Matching window activated."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
