' ローカルユーザーの適用済みグループポリシー結果を照会します。gpupdateは実行しません。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("gpresult.exe /r /scope user")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
