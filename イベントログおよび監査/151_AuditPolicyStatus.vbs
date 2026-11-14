' ファイル概要: auditpolでローカルの監査ポリシーを照会する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 151_AuditPolicyStatus.vbs
' 実行結果: auditpolでローカルの監査ポリシーを照会する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shell, process, line, auditPolPath
Set shell = CreateObject("WScript.Shell")
auditPolPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\auditpol.exe")
Set process = shell.Exec(Chr(34) & auditPolPath & Chr(34) & " /get /category:*")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo "Audit policy query failed with exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
WScript.Echo "Read-only query; audit policy was not changed."
