Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("gpresult.exe /scope computer /r")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo "gpresult returned exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
