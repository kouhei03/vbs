' ファイル概要: Defenderリアルタイム保護と定義バージョン。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 130_DefenderStatus.vbs
' 実行結果: Defenderリアルタイム保護と定義バージョン処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line, commandLine
Set shell = CreateObject("WScript.Shell")
commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-MpComputerStatus | Select-Object RealTimeProtectionEnabled,AntivirusSignatureVersion | Format-List" & Chr(34)
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process = shell.Exec(commandLine)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
