' 指定された単一ホストのPingを1回行います。セグメント走査はしません。
Option Explicit
Dim host, re, shell, runner
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 023_CheckSingleHost.vbs <hostname-or-ip>"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
If Not re.Test(host) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("ping.exe -n 1 -w 2000 " & host)
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit 1
