' 解説: 指定されたウィンドウタイトルの既存ウィンドウにフォーカスします。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, title, activated
Dim shell, title, activated
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 106_ActivateWindowTitle.vbs <window-title>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
title = WScript.Arguments(0)
' 解説: 条件を判定し、必要な分岐を選択します: If Len(title) = 0
If Len(title) = 0 Then WScript.Echo "Window title must not be empty.": WScript.Quit 2
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
activated = shell.AppActivate(title)
' 解説: 条件を判定し、必要な分岐を選択します: If activated
If activated Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Matching window activated."
' 解説: 先行条件が成立しない場合の代替処理へ進みます。
Else
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "No matching window found."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
