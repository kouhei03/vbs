' 指定した単一ホストへのリモートシャットダウンは実行せず、引数の確認だけ行います。
Option Explicit
Dim host, re, message
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 041_RemoteShutdownPreview.vbs <hostname> [message]"
    WScript.Quit 2
End If
host = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
If Not re.Test(host) Then WScript.Quit 2
message = ""
If WScript.Arguments.Count = 2 Then message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
WScript.Echo "Preview only: remote shutdown was not issued."
WScript.Echo "Host=" & host & " Message=" & message
