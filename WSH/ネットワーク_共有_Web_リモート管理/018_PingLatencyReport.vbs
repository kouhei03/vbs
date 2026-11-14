' 概要: 単一ホストに最大5回Pingし、コマンド出力を新規ローカルログへ保存します。
' 実行方法: Usage: cscript //nologo 018_PingLatencyReport.vbs <hostname-or-ip> <new-log-file>
' 処理結果: 単一ホストへの限定Ping結果を新規ログへ保存。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim host, re, outputPath, fso, shell, runner, output, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("ping.exe -n 5 -w 1500 " & host)
Set output = fso.CreateTextFile(outputPath, False, True)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    output.WriteLine line
    WScript.Echo line
Loop
output.Close
If runner.ExitCode <> 0 Then WScript.Quit 1
