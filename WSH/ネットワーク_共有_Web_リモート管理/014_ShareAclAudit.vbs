' 概要: 指定した単一共有パスのACLをicaclsで表示します。権限変更はしません。
' 実行方法: Usage: cscript //nologo 014_ShareAclAudit.vbs <local-or-UNC-path>
' 処理結果: 単一共有パスのACL照会。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim path, re, shell, runner, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 014_ShareAclAudit.vbs <local-or-UNC-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^(?:[A-Za-z]:\\|\\\\)[A-Za-z0-9 _.$-][A-Za-z0-9 _.$\\-]*$"
If Not re.Test(path) Then
    WScript.Echo "Use a simple local path or UNC path."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("icacls.exe """ & path & """")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
