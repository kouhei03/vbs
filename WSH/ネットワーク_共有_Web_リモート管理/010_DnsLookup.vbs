' 指定した単一ホスト名をDNS解決します。逆引きは単一IPアドレスのみ対象です。
Option Explicit
Dim host, re, shell, runner, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 010_DnsLookup.vbs <hostname-or-ip>"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9.-]{1,253}$"
If Not re.Test(host) Or Left(host, 1) = "-" Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("nslookup.exe " & host)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    WScript.Echo line
Loop
If runner.ExitCode <> 0 Then WScript.Quit 1
