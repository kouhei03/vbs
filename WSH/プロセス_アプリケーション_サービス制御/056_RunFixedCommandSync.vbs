' 固定のWindowsバージョン照会コマンドを同期実行し、終了コードを表示します。
Option Explicit
Dim shell, exitCode
Set shell = CreateObject("WScript.Shell")
exitCode = shell.Run("cmd.exe /c ver", 1, True)
WScript.Echo "ExitCode=" & exitCode
If exitCode <> 0 Then WScript.Quit exitCode
