' 概要: normal/debugスイッチに応じてログ詳細度を切り替える例です。
' 実行方法: Usage: cscript //nologo 090_VerboseModeDemo.vbs <normal|debug>
' 処理結果: normal/debugログ切替例。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim mode, verbose
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    ' 取得した値または処理状態をコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 090_VerboseModeDemo.vbs <normal|debug>"
    WScript.Quit 2
End If
' 主処理: normal/debugログ切替例。
mode = LCase(WScript.Arguments(0))
If mode <> "normal" And mode <> "debug" Then WScript.Quit 2
verbose = (mode = "debug")
WScript.Echo "Mode=" & mode
If verbose Then
    WScript.Echo "Timestamp=" & Now
    WScript.Echo "Script=" & WScript.ScriptFullName
    WScript.Echo "ArgumentCount=" & WScript.Arguments.Count
End If
