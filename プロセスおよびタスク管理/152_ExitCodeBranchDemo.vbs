' 解説: 固定された終了コード5のデモコマンドを実行し、終了コードで分岐します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, process, exitCode
Dim shell, process, exitCode
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 固定されたローカル照会コマンドを実行します。
Set process=shell.Exec("cmd.exe /c exit 5")
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdOut.AtEndOfStream
Do While Not process.StdOut.AtEndOfStream
    ' 解説: この文でローカル診断処理を進めます: process.StdOut.ReadLine
    process.StdOut.ReadLine
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 条件が成立している間、繰り返します: Do While Not process.StdErr.AtEndOfStream
Do While Not process.StdErr.AtEndOfStream
    ' 解説: この文でローカル診断処理を進めます: process.StdErr.ReadLine
    process.StdErr.ReadLine
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: 取得値または計算結果を変数に格納します。
exitCode=process.ExitCode
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "ExitCode=" & exitCode
' 解説: 条件を判定し、必要な分岐を選択します: If exitCode=0
If exitCode=0 Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Success branch."
' 解説: 先行条件が成立しない場合の代替処理へ進みます。
Else
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Failure branch; this is an intentional demo exit code."
' 解説: 条件分岐ブロックを終了します。
End If
