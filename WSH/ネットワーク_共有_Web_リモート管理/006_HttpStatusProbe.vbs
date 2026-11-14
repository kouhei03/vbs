' 指定された単一HTTP/HTTPS URLへGETし、ステータスのみ取得します（本文は保存しません）。
Option Explicit
Dim url, re, request
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 006_HttpStatusProbe.vbs <http(s)-url>"
    WScript.Quit 2
End If
url = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^https?://[^ ]+$"
re.IgnoreCase = True
If Not re.Test(url) Then WScript.Quit 2
Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
request.SetTimeouts 3000, 3000, 5000, 5000
request.Open "GET", url, False
request.Send
WScript.Echo "HTTP " & request.Status & " " & request.StatusText
If request.Status < 200 Or request.Status >= 400 Then WScript.Quit 1
