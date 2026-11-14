' 指定環境変数の存在のみを確認します。値の削除は行いません。
Option Explicit
Dim name, re, shell, env
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 156_EnvironmentDeletePreview.vbs <variable-name>"
    WScript.Quit 2
End If
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
