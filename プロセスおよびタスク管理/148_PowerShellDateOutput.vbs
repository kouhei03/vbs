' ファイル概要: 固定PowerShell照会例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 148_PowerShellDateOutput.vbs
' 実行結果: 固定PowerShell照会例処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line
Set shell=CreateObject("WScript.Shell")
Set process=shell.Exec("powershell.exe -NoProfile -NonInteractive -Command Get-Date -Format o")
Do While Not process.StdOut.AtEndOfStream
    line=process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line=process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode<>0 Then WScript.Quit process.ExitCode
