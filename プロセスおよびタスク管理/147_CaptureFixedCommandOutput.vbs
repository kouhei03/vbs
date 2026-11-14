' ファイル概要: 固定コマンドの出力/エラー取得例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 147_CaptureFixedCommandOutput.vbs
' 実行結果: 固定コマンドの出力/エラー取得例処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, outputText, errorText
Set shell=CreateObject("WScript.Shell")
Set process=shell.Exec("cmd.exe /c ver")
outputText=""
errorText=""
Do While Not process.StdOut.AtEndOfStream
    outputText=outputText & process.StdOut.ReadLine & vbCrLf
Loop
Do While Not process.StdErr.AtEndOfStream
    errorText=errorText & process.StdErr.ReadLine & vbCrLf
Loop
WScript.Echo "STDOUT:" & vbCrLf & outputText
WScript.Echo "STDERR:" & vbCrLf & errorText
WScript.Echo "ExitCode=" & process.ExitCode
