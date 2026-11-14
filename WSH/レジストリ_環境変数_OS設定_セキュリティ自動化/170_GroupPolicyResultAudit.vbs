' 概要: ローカルユーザーの適用済みグループポリシー結果を照会します。gpupdateは実行しません。
' 実行方法: cscript //nologo 170_GroupPolicyResultAudit.vbs
' 処理結果: 適用済みユーザーポリシー照会。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("gpresult.exe /r /scope user")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
