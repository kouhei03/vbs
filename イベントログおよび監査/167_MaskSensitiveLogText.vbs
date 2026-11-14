Option Explicit

Dim text, expression
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
expression.Pattern = "(password|passwd|token|secret|api[_-]?key)\s*[:=]\s*[^;\s,]+"
WScript.Echo expression.Replace(text, "$1=[REDACTED]")
