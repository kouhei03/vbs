' 概要: 指定された単一HTTP/HTTPS URLをGETし、応答時間とステータスのみ記録します。
' 実行方法: Usage: cscript //nologo 033_WebResponseTime.vbs <http(s)-url>
' 処理結果: 単一URLの応答時間・HTTP状態確認。処理結果や状態はコンソールに表示されます。
' コード解説: HTTP要求を送り、応答コードや応答内容を確認します。
Option Explicit
Dim url, re, request, started, elapsed
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 033_WebResponseTime.vbs <http(s)-url>"
    WScript.Quit 2
End If
url = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^https?://[^ ]+$"
re.IgnoreCase = True
If Not re.Test(url) Then WScript.Quit 2
' 主処理: 指定された単一URLへHTTP要求を送信します。
Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
request.SetTimeouts 3000, 3000, 5000, 5000
started = Timer
request.Open "GET", url, False
request.Send
elapsed = (Timer - started + 86400) Mod 86400
WScript.Echo "HTTP=" & request.Status & " ElapsedMs=" & CLng(elapsed * 1000)
If elapsed > 3 Then WScript.Echo "WARNING: response exceeded 3 seconds."
