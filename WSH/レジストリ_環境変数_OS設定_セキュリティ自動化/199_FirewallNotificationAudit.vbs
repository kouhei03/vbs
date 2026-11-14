' Firewall通知とプロファイル有効状態を照会します。通知・通信ルールは変更しません。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-NetFirewallProfile | Select-Object Name,Enabled,NotifyOnListen | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
