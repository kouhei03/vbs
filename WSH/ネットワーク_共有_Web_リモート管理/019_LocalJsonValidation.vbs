' ローカルのJSONファイルをPowerShell ConvertFrom-Jsonで構文検証します。
Option Explicit
Dim fso, path, shell, command, runner, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 019_LocalJsonValidation.vbs <local-json-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
command = "$null = Get-Content -Raw -LiteralPath '" & Replace(path, "'", "''") & "' | ConvertFrom-Json"
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & command & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    If line <> "" Then WScript.Echo line
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then
    WScript.Echo "Invalid JSON or PowerShell JSON parser failure."
    WScript.Quit runner.ExitCode
End If
WScript.Echo "JSON syntax parsed successfully."
