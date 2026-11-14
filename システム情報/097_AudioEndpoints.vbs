' ファイル概要: オーディオ入出力エンドポイント状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 097_AudioEndpoints.vbs
' 実行結果: オーディオ入出力エンドポイント状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line, commandLine
Set shell = CreateObject("WScript.Shell")
commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-PnpDevice -Class AudioEndpoint | Select-Object FriendlyName,Status,InstanceId | Format-List" & Chr(34)
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
