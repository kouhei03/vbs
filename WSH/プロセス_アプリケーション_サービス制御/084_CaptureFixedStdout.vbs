' 概要: 固定のverコマンドの標準出力を新規ログへ保存します（任意コマンドは実行しません）。
' 実行方法: Usage: cscript //nologo 084_CaptureFixedStdout.vbs <new-log-file>
' 処理結果: 固定コマンドの標準出力を新規ログへ保存。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, path, shell, runner, output, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 084_CaptureFixedStdout.vbs <new-log-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If fso.FileExists(path) Then
    WScript.Echo "Refusing to overwrite: " & path
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("cmd.exe /c ver")
Set output = fso.CreateTextFile(path, False, True)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    output.WriteLine line
    WScript.Echo line
Loop
output.Close
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
