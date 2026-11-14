' ファイル概要: 更新前の関連プロセス確認。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 091_UpdatePreflight.vbs <process.exe> [process.exe ...]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 更新前の関連プロセス確認処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processName, validator, records, i
If WScript.Arguments.Count < 1 Then WScript.Echo "Usage: cscript //nologo 091_UpdatePreflight.vbs <process.exe> [process.exe ...]": WScript.Quit 2
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 0 To WScript.Arguments.Count - 1
    processName = WScript.Arguments(i)
    If Not validator.Test(processName) Then WScript.Echo "Invalid name: " & processName: WScript.Quit 2
    ' [処理解説] WMIの Win32_Process を照会し、ProcessId を取得します。
    Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    WScript.Echo processName & " running instances=" & records.Count
Next
WScript.Echo "Preflight only; no application was closed."
