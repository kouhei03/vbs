' 固定のverコマンドを有界時間だけ監視します。タイムアウト時もプロセスを強制終了しません。
Option Explicit
Dim seconds, shell, runner, started
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 086_FixedCommandTimeoutAudit.vbs <timeout-seconds-1..30>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 1 Or CDbl(WScript.Arguments(0)) > 30 Then WScript.Quit 2
seconds = CLng(WScript.Arguments(0))
Set shell = CreateObject("WScript.Shell")
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
