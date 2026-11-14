' 概要: 指定された単一ホストのPingを1回行います。セグメント走査はしません。
' 実行方法: Usage: cscript //nologo 023_CheckSingleHost.vbs <hostname-or-ip>
' 処理結果: 単一ホストへの1回Ping。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim host, re, shell, runner
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 023_CheckSingleHost.vbs <hostname-or-ip>"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
If Not re.Test(host) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("ping.exe -n 1 -w 2000 " & host)
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit 1
