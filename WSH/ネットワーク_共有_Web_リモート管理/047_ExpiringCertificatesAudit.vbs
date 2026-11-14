' 概要: ローカルコンピューターのPersonalストアで90日以内に期限を迎える証明書を表示します。
' 実行方法: cscript //nologo 047_ExpiringCertificatesAudit.vbs
' 処理結果: ローカル証明書ストアの90日以内期限切れ監査。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.NotAfter -le (Get-Date).AddDays(90) -and $_.NotAfter -ge (Get-Date) } | Select-Object Subject,Thumbprint,NotAfter | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
