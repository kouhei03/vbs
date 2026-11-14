' 概要: 指定環境変数の存在のみを確認します。値の削除は行いません。
' 実行方法: Usage: cscript //nologo 156_EnvironmentDeletePreview.vbs <variable-name>
' 処理結果: 環境変数の削除候補確認（削除なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellから環境変数を取得し、内容を確認します。
Option Explicit
Dim name, re, shell, env
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 156_EnvironmentDeletePreview.vbs <variable-name>"
    WScript.Quit 2
End If
' 主処理: 環境変数の削除候補確認（削除なし）。
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z_][A-Za-z0-9_]*$"
If Not re.Test(name) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
If env(name) = "" Then
    WScript.Echo "Variable is absent or empty: " & name
Else
    WScript.Echo "Variable exists in this process: " & name
End If
WScript.Echo "Preview only; no environment variable was deleted."
