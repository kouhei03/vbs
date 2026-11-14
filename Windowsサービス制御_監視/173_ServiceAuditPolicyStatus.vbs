' Windows監査ポリシーを照会します。ポリシーは変更しません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: shell, process, line
Dim shell, process, line
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: サービス構成を表示する固定の読み取り専用コマンドを実行します。
Set process = shell.Exec("auditpol.exe /get /category:* ")
' 解説: 条件に従って処理を繰り返します。
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    line = process.StdOut.ReadLine
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo line
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: 条件に従って処理を繰り返します。
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    line = process.StdErr.ReadLine
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "ERROR: " & line
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
