' 概要: ローカルのJSONファイルをPowerShell ConvertFrom-Jsonで構文検証します。
' 実行方法: Usage: cscript //nologo 019_LocalJsonValidation.vbs <local-json-file>
' 処理結果: ローカルJSONファイルの構文検証。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, path, shell, command, runner, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 019_LocalJsonValidation.vbs <local-json-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
command = "$null = Get-Content -Raw -LiteralPath '" & Replace(path, "'", "''") & "' | ConvertFrom-Json"
' 主処理: 外部コマンドの出力を読み取ります。
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
