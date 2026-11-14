' ファイル概要: 指定されたテキストの一部の機密値をマスクする。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 167_MaskSensitiveLogText.vbs <text>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定した文字列内の対象パターンを置換した結果を表示します。すべての機密情報を検出できるわけではありません。

Option Explicit

Dim text, expression
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 167_MaskSensitiveLogText.vbs <text>"
    WScript.Quit 2
End If
text = WScript.Arguments(0)
If Len(text) > 4096 Then
    WScript.Echo "Input is too long (maximum 4096 characters)."
    WScript.Quit 2
End If

Set expression = New RegExp
expression.Global = True
expression.IgnoreCase = True
' [処理解説] 指定したキー名と区切りに一致する値だけをマスクします。すべての機密情報を検出するものではありません。
expression.Pattern = "(password|passwd|token|secret|api[_-]?key)\s*[:=]\s*[^;\s,]+"
WScript.Echo expression.Replace(text, "$1=[REDACTED]")
