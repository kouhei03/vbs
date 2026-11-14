Option Explicit

Dim sourceName, eventId, messageText
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 144_CustomEventWritePreview.vbs <source-name> <event-id> <message>"
    WScript.Quit 2
End If
sourceName = WScript.Arguments(0)
messageText = WScript.Arguments(2)
If Len(sourceName) = 0 Or Len(sourceName) > 128 Then
    WScript.Echo "Source name must contain 1-128 characters."
    WScript.Quit 2
End If
If Not ParseEventId(WScript.Arguments(1), eventId) Then
    WScript.Echo "Event ID must be an integer from 0 to 65535."
    WScript.Quit 2
End If
If Len(messageText) > 2048 Then
    WScript.Echo "Message is too long (maximum 2048 characters)."
    WScript.Quit 2
End If

WScript.Echo "READ-ONLY PREVIEW; no event was written."
WScript.Echo "Source name: " & sourceName
WScript.Echo "Event ID: " & eventId
WScript.Echo "Message length: " & Len(messageText)

Function ParseEventId(text, value)
    Dim i, character
    ParseEventId = False
    If Len(text) = 0 Or Len(text) > 5 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseEventId = (value >= 0 And value <= 65535)
End Function
