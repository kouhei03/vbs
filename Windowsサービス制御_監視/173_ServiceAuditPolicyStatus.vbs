' Windows監査ポリシーを照会します。ポリシーは変更しません。
Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("auditpol.exe /get /category:* ")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
