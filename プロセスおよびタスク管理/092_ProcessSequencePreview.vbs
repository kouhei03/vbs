' ファイル概要: 起動順序のプレビューのみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 092_ProcessSequencePreview.vbs <process.exe> [process.exe ...]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 起動順序のプレビューのみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim validator, i, processName
If WScript.Arguments.Count < 1 Then WScript.Echo "Usage: cscript //nologo 092_ProcessSequencePreview.vbs <process.exe> [process.exe ...]": WScript.Quit 2
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
For i = 0 To WScript.Arguments.Count - 1
    processName = WScript.Arguments(i)
    If Not validator.Test(processName) Then WScript.Echo "Invalid name: " & processName: WScript.Quit 2
    WScript.Echo CStr(i + 1) & ". " & processName
Next
WScript.Echo "Preview only; no program was started."
