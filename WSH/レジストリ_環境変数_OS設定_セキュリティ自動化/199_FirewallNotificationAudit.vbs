' 概要: Firewall通知とプロファイル有効状態を照会します。通知・通信ルールは変更しません。
' 実行方法: cscript //nologo 199_FirewallNotificationAudit.vbs
' 処理結果: Firewall通知状態確認。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-NetFirewallProfile | Select-Object Name,Enabled,NotifyOnListen | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
