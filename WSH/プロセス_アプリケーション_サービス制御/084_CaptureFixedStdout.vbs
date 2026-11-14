' 固定のverコマンドの標準出力を新規ログへ保存します（任意コマンドは実行しません）。
Option Explicit
Dim fso, path, shell, runner, output, line
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
Set runner = shell.Exec("cmd.exe /c ver")
Set output = fso.CreateTextFile(path, False, True)
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    output.WriteLine line
    WScript.Echo line
Loop
output.Close
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
