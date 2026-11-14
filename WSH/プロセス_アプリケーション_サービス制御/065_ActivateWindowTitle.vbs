' 指定タイトルの既存ウィンドウをアクティブにします。SendKeysや最小化は行いません。
Option Explicit
Dim title, shell
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 065_ActivateWindowTitle.vbs <window-title>"
    WScript.Quit 2
End If
title = WScript.Arguments(0)
If title = "" Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
If shell.AppActivate(title) Then
    WScript.Echo "Activated matching window."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
