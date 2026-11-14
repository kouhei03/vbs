' 概要: シャットダウン操作は実行せず、指定されたカウントダウン条件だけ確認します。
' 実行方法: Usage: cscript //nologo 087_ShutdownPreviewOnly.vbs <requested-delay-seconds>
' 処理結果: シャットダウン条件の確認のみ。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim seconds
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    ' 取得した値または処理状態をコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 087_ShutdownPreviewOnly.vbs <requested-delay-seconds>"
    WScript.Quit 2
End If
' 主処理: シャットダウン条件の確認のみ。
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < 0 Or CDbl(WScript.Arguments(0)) > 86400 Then WScript.Quit 2
seconds = CLng(WScript.Arguments(0))
WScript.Echo "Preview only: a shutdown countdown of " & seconds & " seconds was requested."
WScript.Echo "No shutdown or restart command was issued."
