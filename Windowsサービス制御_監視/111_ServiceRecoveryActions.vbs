' ファイル概要: サービス回復アクション設定の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 111_ServiceRecoveryActions.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス回復アクション設定の読み取り処理の結果やエラーをコンソールに表示します。

' サービスの回復アクション設定を照会します。sc.exeのquery系コマンドのみ実行します。
Option Explicit
Dim serviceName, validator
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 111_ServiceRecoveryActions.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    WScript.Quit 2
End If
Dim shell, process, line, commandLine
Set shell = CreateObject("WScript.Shell")
commandLine = "sc.exe qfailure " & serviceName
Set process = shell.Exec(commandLine)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
