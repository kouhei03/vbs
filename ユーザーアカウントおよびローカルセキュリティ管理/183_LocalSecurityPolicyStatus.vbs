Option Explicit

Dim shell, process, output, errorOutput, auditPolPath
Set shell = CreateObject("WScript.Shell")
auditPolPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\auditpol.exe")
Set process = shell.Exec(Chr(34) & auditPolPath & Chr(34) & " /get /category:*")
output = process.StdOut.ReadAll
errorOutput = process.StdErr.ReadAll
If Len(output) > 0 Then WScript.Echo output
If process.ExitCode <> 0 Then
    If Len(errorOutput) > 0 Then WScript.Echo errorOutput
    WScript.Echo "Audit policy query failed with exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
WScript.Echo "Read-only query; no local security policy was changed."
