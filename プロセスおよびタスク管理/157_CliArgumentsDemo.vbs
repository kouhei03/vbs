' 解説: WScript.Argumentsの個数と各引数を表示する、ローカルCLIの安全な例です。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: i
Dim i
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Argument count=" & WScript.Arguments.Count
' 解説: 指定範囲の要素を順番に処理します: For i=0 To WScript.Arguments.Count-1
For i=0 To WScript.Arguments.Count-1
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Argument " & i & "=" & WScript.Arguments(i)
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
