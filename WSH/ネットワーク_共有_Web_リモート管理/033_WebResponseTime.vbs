' 指定された単一HTTP/HTTPS URLをGETし、応答時間とステータスのみ記録します。
Option Explicit
Dim url, re, request, started, elapsed
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 033_WebResponseTime.vbs <http(s)-url>"
    WScript.Quit 2
End If
url = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^https?://[^ ]+$"
re.IgnoreCase = True
If Not re.Test(url) Then WScript.Quit 2
Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
request.SetTimeouts 3000, 3000, 5000, 5000
started = Timer
request.Open "GET", url, False
request.Send
elapsed = (Timer - started + 86400) Mod 86400
WScript.Echo "HTTP=" & request.Status & " ElapsedMs=" & CLng(elapsed * 1000)
If elapsed > 3 Then WScript.Echo "WARNING: response exceeded 3 seconds."
