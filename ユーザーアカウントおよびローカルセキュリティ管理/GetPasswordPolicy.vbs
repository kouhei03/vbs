Option Explicit

Dim shell, command, standardOutput, standardError
Set shell = CreateObject("WScript.Shell")

' This fixed command displays local account policy and does not change it.
Set command = shell.Exec("net.exe accounts")
standardOutput = command.StdOut.ReadAll
standardError = command.StdErr.ReadAll

If Len(standardOutput) > 0 Then WScript.Echo standardOutput
If command.ExitCode <> 0 Then
    If Len(standardError) > 0 Then WScript.Echo standardError
    WScript.Echo "net accounts failed with exit code " & command.ExitCode
    WScript.Quit command.ExitCode
End If
