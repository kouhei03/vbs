' 解説: 起動予定の実行ファイル名を検証し、指定順を表示するだけのシーケンスプレビューです。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: validator, i, processName
Dim validator, i, processName
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count < 1
If WScript.Arguments.Count < 1 Then WScript.Echo "Usage: cscript //nologo 092_ProcessSequencePreview.vbs <process.exe> [process.exe ...]": WScript.Quit 2
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 指定範囲の要素を順番に処理します: For i = 0 To WScript.Arguments.Count - 1
For i = 0 To WScript.Arguments.Count - 1
    ' 解説: 取得値または計算結果を変数に格納します。
    processName = WScript.Arguments(i)
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName)
    If Not validator.Test(processName) Then WScript.Echo "Invalid name: " & processName: WScript.Quit 2
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo CStr(i + 1) & ". " & processName
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Preview only; no program was started."
