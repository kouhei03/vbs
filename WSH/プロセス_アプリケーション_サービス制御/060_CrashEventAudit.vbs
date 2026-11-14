' 概要: 直近のApplication Errorイベントを表示します。再起動やダンプ取得は行いません。
' 実行方法: cscript //nologo 060_CrashEventAudit.vbs
' 処理結果: Application Errorイベント監査（再起動なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000} -MaxEvents 10 | Select-Object TimeCreated,Id,ProviderName,Message | Format-List" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
