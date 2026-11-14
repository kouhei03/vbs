Option Explicit

Dim userName
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 175_PasswordChangePreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
If Len(userName) = 0 Or Len(userName) > 256 Then
    WScript.Echo "User name length is invalid."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no password was changed."
WScript.Echo "Target local account: " & userName
WScript.Echo "Do not pass passwords on a command line."
