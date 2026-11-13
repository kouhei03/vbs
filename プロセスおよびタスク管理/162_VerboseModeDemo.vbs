' 解説: verbose引数で診断メッセージ量を切り替える例です。既定では簡潔な出力だけを行います。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: verbose, i
Dim verbose, i
' 解説: 取得値または計算結果を変数に格納します。
verbose=False
' 解説: 指定範囲の要素を順番に処理します: For i=0 To WScript.Arguments.Count-1
For i=0 To WScript.Arguments.Count-1
    ' 解説: 条件を判定し、必要な分岐を選択します: If LCase(WScript.Arguments(i))="/verbose"
    If LCase(WScript.Arguments(i))="/verbose" Then verbose=True
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Process diagnostics started."
' 解説: 条件を判定し、必要な分岐を選択します: If verbose
If verbose Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Host=" & CreateObject("WScript.Network").ComputerName
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Timestamp=" & Now
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Arguments=" & WScript.Arguments.Count
' 解説: 条件分岐ブロックを終了します。
End If
