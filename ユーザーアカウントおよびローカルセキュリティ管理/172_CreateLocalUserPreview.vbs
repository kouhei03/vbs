Option Explicit

Dim userName, expression
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 172_CreateLocalUserPreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9._-]{0,19}$"
If Not expression.Test(userName) Then
    WScript.Echo "User name must be 1-20 safe ASCII characters."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no account was created."
WScript.Echo "Proposed local account name: " & userName
WScript.Echo "No password was accepted or displayed."
