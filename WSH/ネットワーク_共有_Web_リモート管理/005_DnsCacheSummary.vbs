' DNSリゾルバーキャッシュを表示する固定コマンド例です。消去はしません。
Option Explicit
Dim shell, runner, line, count
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("ipconfig.exe /displaydns")
count = 0
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    If InStr(1, line, "Record Name", 1) > 0 Then count = count + 1
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
WScript.Echo "DNS cached-record sections=" & count & " (record details omitted)."
