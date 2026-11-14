' ファイル概要: 監査ポリシー状態の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 173_ServiceAuditPolicyStatus.vbs
' 実行結果: 監査ポリシー状態の読み取り処理の結果やエラーをコンソールに表示します。

' Windows監査ポリシーを照会します。ポリシーは変更しません。
Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("auditpol.exe /get /category:* ")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
