' 解説: 固定された読み取り専用PowerShell式で日時文字列を取得します。任意コードは受け取りません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, process, line
Dim shell, process, line
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 固定されたローカル照会コマンドを実行します。
Set process=shell.Exec("powershell.exe -NoProfile -NonInteractive -Command Get-Date -Format o")
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
' 解説: 条件を判定し、必要な分岐を選択します: If process.ExitCode<>0
If process.ExitCode<>0 Then WScript.Quit process.ExitCode
