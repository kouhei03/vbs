' 固定のWindowsバージョン照会を可視コンソールで実行します。
Option Explicit
Dim shell, exitCode
Set shell = CreateObject("WScript.Shell")
exitCode = shell.Run("cmd.exe /c ver", 1, True)
WScript.Echo "Visible fixed command exit code=" & exitCode
If exitCode <> 0 Then WScript.Quit exitCode
