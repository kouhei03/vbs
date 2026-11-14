Option Explicit

Dim shell, process, output, errorOutput, quserPath
Set shell = CreateObject("WScript.Shell")
quserPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\quser.exe")
Set process = shell.Exec(Chr(34) & quserPath & Chr(34))
output = process.StdOut.ReadAll
errorOutput = process.StdErr.ReadAll
If Len(output) > 0 Then WScript.Echo output
If process.ExitCode <> 0 Then
    If Len(errorOutput) > 0 Then WScript.Echo errorOutput
    WScript.Echo "Session query failed with exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
WScript.Echo "Session inventory only; no user was logged off."
