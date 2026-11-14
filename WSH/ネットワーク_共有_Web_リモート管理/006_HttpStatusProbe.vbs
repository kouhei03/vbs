' 概要: 指定された単一HTTP/HTTPS URLへGETし、ステータスのみ取得します（本文は保存しません）。
' 実行方法: Usage: cscript //nologo 006_HttpStatusProbe.vbs <http(s)-url>
' 処理結果: 単一URLへのHTTP GETとステータス確認。処理結果や状態はコンソールに表示されます。
' コード解説: HTTP要求を送り、応答コードや応答内容を確認します。
Option Explicit
Dim url, re, request
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 006_HttpStatusProbe.vbs <http(s)-url>"
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
request.Open "GET", url, False
request.Send
WScript.Echo "HTTP " & request.Status & " " & request.StatusText
If request.Status < 200 Or request.Status >= 400 Then WScript.Quit 1
