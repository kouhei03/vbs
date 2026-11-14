' ファイル概要: 重複起動の検出のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 087_DuplicateProcessCheck.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 重複起動の検出のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, processName, validator
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 087_DuplicateProcessCheck.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、ProcessId を取得します。
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
WScript.Echo "Running instances: " & records.Count
If records.Count > 1 Then WScript.Echo "Duplicate instances detected; no process was blocked or terminated."
