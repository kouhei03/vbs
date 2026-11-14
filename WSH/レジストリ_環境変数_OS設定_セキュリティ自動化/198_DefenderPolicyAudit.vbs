' Microsoft Defenderのクラウド保護・サンプル送信設定を照会します。保護機能は変更しません。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-MpPreference | Select-Object MAPSReporting,SubmitSamplesConsent | Format-List" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
