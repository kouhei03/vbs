Option Explicit
Dim shell, process, startedAt, elapsed, line
Set shell=CreateObject("WScript.Shell")
startedAt=Timer
Set process=shell.Exec("cmd.exe /c ver")
Do While Not process.StdOut.AtEndOfStream
    line=process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line=process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
elapsed=Timer-startedAt
If elapsed<0 Then elapsed=elapsed+86400
WScript.Echo "Elapsed seconds=" & FormatNumber(elapsed,3)
