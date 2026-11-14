' 指定した単一ホストへ最大3回だけPingします。ネットワーク範囲の走査はしません。
Option Explicit
Dim host, re, shell, runner, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 003_PingSingleHost.vbs <hostname-or-ip>"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9.-]{1,253}$"
If Not re.Test(host) Or Left(host, 1) = "-" Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("ping.exe -n 3 -w 1500 " & host)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    WScript.Echo line
Loop
If runner.ExitCode <> 0 Then WScript.Quit 1
