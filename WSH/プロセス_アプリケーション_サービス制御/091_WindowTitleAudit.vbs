' 指定タイトルの既存ウィンドウがあるか確認します。キー送信は行いません。
Option Explicit
Dim title, shell
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 091_WindowTitleAudit.vbs <window-title>"
    WScript.Quit 2
End If
title = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
If shell.AppActivate(title) Then
    WScript.Echo "Matching window exists (it may have been activated)."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
