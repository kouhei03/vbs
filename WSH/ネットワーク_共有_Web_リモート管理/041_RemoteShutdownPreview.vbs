' 概要: 指定した単一ホストへのリモートシャットダウンは実行せず、引数の確認だけ行います。
' 実行方法: Usage: cscript //nologo 041_RemoteShutdownPreview.vbs <hostname> [message]
' 処理結果: リモート停止コマンドを実行しないプレビュー。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim host, re, message
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 041_RemoteShutdownPreview.vbs <hostname> [message]"
    WScript.Quit 2
End If
' 主処理: リモート停止コマンドを実行しないプレビュー。
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
If Not re.Test(host) Then WScript.Quit 2
message = ""
If WScript.Arguments.Count = 2 Then message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
WScript.Echo "Preview only: remote shutdown was not issued."
WScript.Echo "Host=" & host & " Message=" & message
