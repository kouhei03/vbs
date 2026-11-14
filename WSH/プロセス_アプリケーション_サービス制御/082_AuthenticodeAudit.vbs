' PowerShellの固定照会で指定ファイルのAuthenticode状態を表示します。
Option Explicit
Dim fso, path, shell, command, runner, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 082_AuthenticodeAudit.vbs <file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
command = "Get-AuthenticodeSignature -LiteralPath '" & Replace(path, "'", "''") & "' | Select-Object Status,@{Name='Publisher';Expression={$_.SignerCertificate.Subject}} | Format-List"
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & command & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
