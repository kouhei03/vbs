' 解説: 固定のcmd.exe verコマンドの実行時間を計測します。外部データは送信しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, process, startedAt, elapsed, line
Dim shell, process, startedAt, elapsed, line
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
startedAt=Timer
' 解説: 固定されたローカル照会コマンドを実行します。
Set process=shell.Exec("cmd.exe /c ver")
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdOut.AtEndOfStream
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    line=process.StdOut.ReadLine
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo line
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdErr.AtEndOfStream
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    line=process.StdErr.ReadLine
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "ERROR: " & line
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 取得値または計算結果を変数に格納します。
elapsed=Timer-startedAt
' 解説: 条件を判定し、必要な分岐を選択します: If elapsed<0
If elapsed<0 Then elapsed=elapsed+86400
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Elapsed seconds=" & FormatNumber(elapsed,3)
