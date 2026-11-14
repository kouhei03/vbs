Option Explicit

Dim shell, process, output, errorOutput, netPath
Set shell = CreateObject("WScript.Shell")
netPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\net.exe")
Set process = shell.Exec(Chr(34) & netPath & Chr(34) & " accounts")
output = process.StdOut.ReadAll
errorOutput = process.StdErr.ReadAll
If Len(output) > 0 Then WScript.Echo output
If process.ExitCode <> 0 Then
    If Len(errorOutput) > 0 Then WScript.Echo errorOutput
    WScript.Echo "Password policy query failed with exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
WScript.Echo "Read-only query; account policy was not changed."
