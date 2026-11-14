Option Explicit

Dim groupName, expression
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 180_CreateLocalGroupPreview.vbs <group-name>"
    WScript.Quit 2
End If
groupName = WScript.Arguments(0)
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9._-]{0,47}$"
If Not expression.Test(groupName) Then
    WScript.Echo "Group name must be 1-48 safe ASCII characters."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no local group was created."
WScript.Echo "Proposed group name: " & groupName
