' /verbose引数の解釈例です。サービス構成や状態は変更しません。
Option Explicit
Dim debugMode, argument
 debugMode = False
For Each argument In WScript.Arguments
    If LCase(argument) = "/verbose" Then debugMode = True
Next
If debugMode Then
    WScript.Echo "Verbose mode enabled for this local read-only demo."
Else
    WScript.Echo "Normal mode. Pass /verbose for an extra diagnostic line."
End If
WScript.Echo "No service operation was performed."
