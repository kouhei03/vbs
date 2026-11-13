' 解説: 固定されたcmd.exe verコマンドの標準出力/標準エラーを読み取る安全な例です。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, process, outputText, errorText
Dim shell, process, outputText, errorText
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 固定されたローカル照会コマンドを実行します。
Set process=shell.Exec("cmd.exe /c ver")
' 解説: 取得値または計算結果を変数に格納します。
outputText=""
' 解説: 取得値または計算結果を変数に格納します。
errorText=""
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdOut.AtEndOfStream
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    outputText=outputText & process.StdOut.ReadLine & vbCrLf
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdErr.AtEndOfStream
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    errorText=errorText & process.StdErr.ReadLine & vbCrLf
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "STDOUT:" & vbCrLf & outputText
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "STDERR:" & vbCrLf & errorText
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "ExitCode=" & process.ExitCode
