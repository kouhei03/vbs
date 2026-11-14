' 長い処理を想定し、上限付きループでWScript.Sleepを使う例です。
Option Explicit
Dim iterations, intervalMs, i
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 088_BoundedSleepDemo.vbs <iterations-1..20> <sleep-ms-10..5000>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 1 Or CDbl(WScript.Arguments(0)) > 20 Or _
    CDbl(WScript.Arguments(1)) < 10 Or CDbl(WScript.Arguments(1)) > 5000 Then WScript.Quit 2
iterations = CLng(WScript.Arguments(0))
intervalMs = CLng(WScript.Arguments(1))
For i = 1 To iterations
    WScript.Echo "Sample " & i & " of " & iterations
    If i < iterations Then WScript.Sleep intervalMs
Next
