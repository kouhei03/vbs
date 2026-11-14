' 直近のApplication Errorイベントを表示します。再起動やダンプ取得は行いません。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000} -MaxEvents 10 | Select-Object TimeCreated,Id,ProviderName,Message | Format-List" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
