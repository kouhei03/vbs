' 概要: Application Errorイベントの概要を読み取ります。ダンプやメモリ内容は取得しません。
' 実行方法: cscript //nologo 074_CrashEventMetadata.vbs
' 処理結果: クラッシュイベントのメタデータ照会（ダンプなし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & _
    "Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000} -MaxEvents 10 | Select-Object TimeCreated,ProviderName,Id | Format-Table -AutoSize" & Chr(34))
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
