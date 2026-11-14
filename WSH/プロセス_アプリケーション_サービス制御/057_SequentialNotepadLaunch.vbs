' 概要: 固定アプリNotepadを1～5回、指定間隔を空けて順に起動する例です。
' 実行方法: Usage: cscript //nologo 057_SequentialNotepadLaunch.vbs <count-1..5> <delay-seconds-1..30>
' 処理結果: Notepadの有界な順次起動。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellを使ってWindowsの設定や状態を照会します。
Option Explicit
Dim count, delaySeconds, i, shell
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 057_SequentialNotepadLaunch.vbs <count-1..5> <delay-seconds-1..30>"
    WScript.Quit 2
End If
' 主処理: Notepadの有界な順次起動。
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
