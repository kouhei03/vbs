' シャットダウン操作は実行せず、指定されたカウントダウン条件だけ確認します。
Option Explicit
Dim seconds
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 087_ShutdownPreviewOnly.vbs <requested-delay-seconds>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 0 Or CDbl(WScript.Arguments(0)) > 86400 Then WScript.Quit 2
seconds = CLng(WScript.Arguments(0))
WScript.Echo "Preview only: a shutdown countdown of " & seconds & " seconds was requested."
WScript.Echo "No shutdown or restart command was issued."
