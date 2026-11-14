' 概要: Microsoft Defenderのクラウド保護・サンプル送信設定を照会します。保護機能は変更しません。
' 実行方法: cscript //nologo 198_DefenderPolicyAudit.vbs
' 処理結果: Defenderクラウド保護等の状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-MpPreference | Select-Object MAPSReporting,SubmitSamplesConsent | Format-List" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
