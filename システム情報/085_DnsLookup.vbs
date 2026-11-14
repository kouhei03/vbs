' ファイル概要: ホスト名DNSルックアップ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 085_DnsLookup.vbs <hostname>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ホスト名DNSルックアップ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim hostname, validator, shell, process, line
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 085_DnsLookup.vbs <hostname>"
    WScript.Quit 2
End If
hostname = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]*$"
If Not validator.Test(hostname) Then
    WScript.Echo "Invalid host name."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process = shell.Exec("nslookup.exe " & hostname)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
