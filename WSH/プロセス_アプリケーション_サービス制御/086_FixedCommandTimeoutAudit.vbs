' 概要: 固定のverコマンドを有界時間だけ監視します。タイムアウト時もプロセスを強制終了しません。
' 実行方法: Usage: cscript //nologo 086_FixedCommandTimeoutAudit.vbs <timeout-seconds-1..30>
' 処理結果: 固定コマンドの期限監視（強制終了なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim seconds, shell, runner, started
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 086_FixedCommandTimeoutAudit.vbs <timeout-seconds-1..30>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 1 Or CDbl(WScript.Arguments(0)) > 30 Then WScript.Quit 2
seconds = CLng(WScript.Arguments(0))
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("cmd.exe /c ver")
started = Timer
Do While runner.Status = 0 And (Timer - started + 86400) Mod 86400 < seconds
    WScript.Sleep 100
Loop
If runner.Status = 0 Then
    WScript.Echo "Completed. ExitCode=" & runner.ExitCode
Else
    WScript.Echo "Timeout observed; child was not terminated."
End If
