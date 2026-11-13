' 解説: tasklistの詳細表示から応答なしプロセス候補を抽出します。自動終了はしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, process, line, found
Dim shell, process, line, found
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 固定されたローカル照会コマンドを実行します。
Set process = shell.Exec("tasklist.exe /v /fo csv")
' 解説: 取得値または計算結果を変数に格納します。
found = False
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdOut.AtEndOfStream
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    line = process.StdOut.ReadLine
    ' 解説: 条件を判定し、必要な分岐を選択します: If InStr(1, line, "Not Responding", vbTextCompare) > 0 Or InStr(line, "応答なし") > 0
    If InStr(1, line, "Not Responding", vbTextCompare) > 0 Or InStr(line, "応答なし") > 0 Then
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo line
        ' 解説: 取得値または計算結果を変数に格納します。
        found = True
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 条件を判定し、必要な分岐を選択します: If Not found
If Not found Then WScript.Echo "No not-responding process row found."
' 解説: 条件を判定し、必要な分岐を選択します: If process.ExitCode <> 0
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
