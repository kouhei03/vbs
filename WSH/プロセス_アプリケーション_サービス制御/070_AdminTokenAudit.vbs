' PowerShellの固定照会で現在のプロセスが管理者ロールか確認します（昇格はしません）。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    line = Trim(runner.StdOut.ReadLine)
    If line <> "" Then WScript.Echo "AdministratorRole=" & line
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
