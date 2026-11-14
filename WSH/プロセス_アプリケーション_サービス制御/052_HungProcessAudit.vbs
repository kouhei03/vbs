' 応答なしプロセスを一覧表示します。プロセス終了は行いません。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
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
