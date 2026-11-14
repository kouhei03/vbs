' 概要: 長い処理を想定し、上限付きループでWScript.Sleepを使う例です。
' 実行方法: Usage: cscript //nologo 088_BoundedSleepDemo.vbs <iterations-1..20> <sleep-ms-10..5000>
' 処理結果: 上限付きSleepポーリング例。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim iterations, intervalMs, i
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    ' 取得した値または処理状態をコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 088_BoundedSleepDemo.vbs <iterations-1..20> <sleep-ms-10..5000>"
    WScript.Quit 2
End If
' 主処理: 上限付きSleepポーリング例。
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 1 Or CDbl(WScript.Arguments(0)) > 20 Or _
    CDbl(WScript.Arguments(1)) < 10 Or CDbl(WScript.Arguments(1)) > 5000 Then WScript.Quit 2
iterations = CLng(WScript.Arguments(0))
intervalMs = CLng(WScript.Arguments(1))
For i = 1 To iterations
    WScript.Echo "Sample " & i & " of " & iterations
    If i < iterations Then WScript.Sleep intervalMs
Next
