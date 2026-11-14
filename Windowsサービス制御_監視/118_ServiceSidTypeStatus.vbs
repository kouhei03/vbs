' サービスSIDの型をsc.exeで照会します。
Option Explicit
Dim serviceName, validator
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 118_ServiceSidTypeStatus.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    WScript.Quit 2
End If
Dim shell, process, line, commandLine
Set shell = CreateObject("WScript.Shell")
commandLine = "sc.exe qsidtype " & serviceName
Set process = shell.Exec(commandLine)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
