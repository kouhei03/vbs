Option Explicit

Dim computerName, pattern
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 150_RemoteLogAccessPreview.vbs <computer-name>"
    WScript.Quit 2
End If
computerName = WScript.Arguments(0)
Set pattern = New RegExp
pattern.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
pattern.IgnoreCase = True
If Not pattern.Test(computerName) Then
    WScript.Echo "Computer name contains unsupported characters."
    WScript.Quit 2
End If
WScript.Echo "Target name accepted for planning: " & computerName
WScript.Echo "No network connection or remote event-log query was made."
