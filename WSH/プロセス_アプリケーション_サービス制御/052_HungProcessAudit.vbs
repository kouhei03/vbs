' 概要: 応答なしプロセスを一覧表示します。プロセス終了は行いません。
' 実行方法: cscript //nologo 052_HungProcessAudit.vbs
' 処理結果: 応答なしプロセスの監査（終了なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-Process | Where-Object { -not $_.Responding } | Select-Object ProcessName,Id | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
WScript.Echo "Audit only; no process was terminated."
