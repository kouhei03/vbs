' 固定の診断文字列をstderrへ出し、標準エラー検知例を表示します。通知送信は行いません。
Option Explicit
Dim shell, runner, line, hadError
Set shell = CreateObject("WScript.Shell")
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
