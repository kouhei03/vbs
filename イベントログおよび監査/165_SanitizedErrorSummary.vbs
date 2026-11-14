' ファイル概要: 指定されたエラー情報の一部の機密値をマスクして表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 165_SanitizedErrorSummary.vbs <number> <source> <description>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定されたエラー情報の一部の機密値をマスクして表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim errorNumber, errorSource, errorDescription, integerPattern
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 165_SanitizedErrorSummary.vbs <number> <source> <description>"
    WScript.Quit 2
End If
Set integerPattern = New RegExp
integerPattern.Pattern = "^-?[0-9]{1,10}$"
If Not integerPattern.Test(WScript.Arguments(0)) Then
    WScript.Echo "Error number must be a signed integer."
    WScript.Quit 2
End If
If CDbl(WScript.Arguments(0)) < -2147483648 Or _
   CDbl(WScript.Arguments(0)) > 2147483647 Then
    WScript.Echo "Error number must be an integer in the signed 32-bit range."
    WScript.Quit 2
End If
errorNumber = CLng(WScript.Arguments(0))
errorSource = Left(WScript.Arguments(1), 128)
errorDescription = WScript.Arguments(2)
If Len(errorDescription) > 2048 Then errorDescription = Left(errorDescription, 2048)
errorDescription = Redact(errorDescription)

WScript.Echo "Error number: " & errorNumber
WScript.Echo "Source: " & Redact(errorSource)
WScript.Echo "Sanitized description: " & errorDescription

Function Redact(text)
    Dim expression
    Set expression = New RegExp
    expression.Global = True
    expression.IgnoreCase = True
    expression.Pattern = "(password|passwd|token|secret|api[_-]?key)\s*[:=]\s*[^;\s,]+"
    Redact = expression.Replace(text, "$1=[REDACTED]")
End Function
