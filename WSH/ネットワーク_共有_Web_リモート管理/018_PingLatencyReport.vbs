' 単一ホストに最大5回Pingし、コマンド出力を新規ローカルログへ保存します。
Option Explicit
Dim host, re, outputPath, fso, shell, runner, output, line
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 018_PingLatencyReport.vbs <hostname-or-ip> <new-log-file>"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
If Not re.Test(host) Then WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If fso.FileExists(outputPath) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("ping.exe -n 5 -w 1500 " & host)
Set output = fso.CreateTextFile(outputPath, False, True)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    output.WriteLine line
    WScript.Echo line
Loop
output.Close
If runner.ExitCode <> 0 Then WScript.Quit 1
