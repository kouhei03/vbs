' 現在のローカルアカウントロックアウトポリシーをnet accountsで表示します（変更なし）。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("net.exe accounts")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
