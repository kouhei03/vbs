' ファイル概要: 直近の電源復帰理由。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 115_LastWakeReason.vbs
' 実行結果: 直近の電源復帰理由処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process = shell.Exec("powercfg.exe /lastwake")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
