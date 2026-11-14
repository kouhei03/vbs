Option Explicit
Dim hostname, validator, shell, process, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 085_DnsLookup.vbs <hostname>"
    WScript.Quit 2
End If
hostname = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]*$"
If Not validator.Test(hostname) Then
    WScript.Echo "Invalid host name."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("nslookup.exe " & hostname)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
