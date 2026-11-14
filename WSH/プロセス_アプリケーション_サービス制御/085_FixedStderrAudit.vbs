' 概要: 固定の診断文字列をstderrへ出し、標準エラー検知例を表示します。通知送信は行いません。
' 実行方法: cscript //nologo 085_FixedStderrAudit.vbs
' 処理結果: 固定デモの標準エラー検知（通知送信なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line, hadError
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("cmd.exe /c echo diagnostic-stderr-demo 1>&2")
hadError = False
Do While Not runner.StdErr.AtEndOfStream
    line = runner.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
    hadError = True
Loop
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
If hadError Then WScript.Echo "Stderr was detected; no external notification was sent."
