' 固定アプリNotepadを1～5回、指定間隔を空けて順に起動する例です。
Option Explicit
Dim count, delaySeconds, i, shell
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 057_SequentialNotepadLaunch.vbs <count-1..5> <delay-seconds-1..30>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 1 Or CDbl(WScript.Arguments(0)) > 5 Or _
    CDbl(WScript.Arguments(1)) < 1 Or CDbl(WScript.Arguments(1)) > 30 Then WScript.Quit 2
count = CLng(WScript.Arguments(0))
delaySeconds = CLng(WScript.Arguments(1))
Set shell = CreateObject("WScript.Shell")
For i = 1 To count
    shell.Run "notepad.exe", 1, False
    WScript.Echo "Launched fixed app " & i & " of " & count
    If i < count Then WScript.Sleep delaySeconds * 1000
Next
