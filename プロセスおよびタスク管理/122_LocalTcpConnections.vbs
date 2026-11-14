' ファイル概要: ローカルTCP接続の一覧。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 122_LocalTcpConnections.vbs
' 実行結果: ローカルTCP接続の一覧処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-NetTCPConnection | Select-Object LocalAddress,LocalPort,State,OwningProcess | Format-Table -AutoSize" & Chr(34))
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
