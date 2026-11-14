' ファイル概要: タスクスケジューラ状態照会のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 151_ScheduledTaskStatus.vbs
' 実行結果: タスクスケジューラ状態照会のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line
Set shell=CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process=shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-ScheduledTask | Select-Object TaskPath,TaskName,State | Format-Table -AutoSize" & Chr(34))
Do While Not process.StdOut.AtEndOfStream
    line=process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line=process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode<>0 Then WScript.Quit process.ExitCode
