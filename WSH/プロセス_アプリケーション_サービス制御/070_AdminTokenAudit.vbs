' 概要: PowerShellの固定照会で現在のプロセスが管理者ロールか確認します（昇格はしません）。
' 実行方法: cscript //nologo 070_AdminTokenAudit.vbs
' 処理結果: 管理者ロール照会（昇格なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    line = Trim(runner.StdOut.ReadLine)
    If line <> "" Then WScript.Echo "AdministratorRole=" & line
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
