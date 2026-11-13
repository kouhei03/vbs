' 解説: スリープ/休止状態対応を組み込みWindowsコマンドで照会し、結果をローカル表示します。
Option Explicit
' 解説: コマンド実行、標準出力、および標準エラーの読み取りに使う変数を宣言します。
Dim shell, process, line
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 読み取り専用の診断コマンドを実行します。
Set process = shell.Exec("powercfg.exe /a")
' 解説: コマンドの標準出力を一行ずつ表示します。
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdOut.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 標準エラーも明示し、失敗内容を隠しません。
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdErr.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "ERROR: " & line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: process.ExitCode <> 0
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
