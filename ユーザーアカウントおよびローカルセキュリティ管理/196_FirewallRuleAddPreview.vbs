Option Explicit

Dim ruleName, protocol, port
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 196_FirewallRuleAddPreview.vbs <name> <TCP|UDP> <port 1-65535>"
    WScript.Quit 2
End If
ruleName = WScript.Arguments(0)
protocol = UCase(WScript.Arguments(1))
If Len(ruleName) = 0 Or Len(ruleName) > 128 Then
    WScript.Echo "Rule name length is invalid."
    WScript.Quit 2
End If
If protocol <> "TCP" And protocol <> "UDP" Then
    WScript.Echo "Protocol must be TCP or UDP."
    WScript.Quit 2
End If
If Not ParsePort(WScript.Arguments(2), port) Then
    WScript.Echo "Port must be an integer from 1 to 65535."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY firewall preview; no rule was created."
WScript.Echo "Rule name: " & ruleName
WScript.Echo "Protocol/port: " & protocol & "/" & port
WScript.Echo "Direction, scope, profile, and business approval still require review."

Function ParsePort(text, value)
    Dim i, character
    ParsePort = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParsePort = (value >= 1 And value <= 65535)
End Function
