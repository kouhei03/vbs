' ローカルコンピューターのPersonalストアで90日以内に期限を迎える証明書を表示します。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.NotAfter -le (Get-Date).AddDays(90) -and $_.NotAfter -ge (Get-Date) } | Select-Object Subject,Thumbprint,NotAfter | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
