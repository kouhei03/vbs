Option Explicit
Dim shell, title, activated
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 106_ActivateWindowTitle.vbs <window-title>": WScript.Quit 2
title = WScript.Arguments(0)
If Len(title) = 0 Then WScript.Echo "Window title must not be empty.": WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
activated = shell.AppActivate(title)
If activated Then
    WScript.Echo "Matching window activated."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
