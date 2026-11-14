Option Explicit

Dim folderPath, ageDays, fso, folder, fileItem, cutoff, oldCount, oldBytes
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 166_LogRotationPreflight.vbs <folder> <age-days 1-3650>"
    WScript.Quit 2
End If
folderPath = WScript.Arguments(0)
If Not ParseDays(WScript.Arguments(1), ageDays) Then
    WScript.Echo "Age must be a whole number from 1 to 3650."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(folderPath) Then
    WScript.Echo "Folder does not exist or is unavailable."
    WScript.Quit 1
End If
Set folder = fso.GetFolder(folderPath)
cutoff = DateAdd("d", -ageDays, Now)
oldCount = 0
oldBytes = 0
For Each fileItem In folder.Files
    If fileItem.DateLastModified < cutoff Then
        oldCount = oldCount + 1
        oldBytes = oldBytes + fileItem.Size
    End If
Next
WScript.Echo "Files older than " & ageDays & " days: " & oldCount
WScript.Echo "Total bytes: " & oldBytes
WScript.Echo "Preflight only; no files were compressed or deleted."

Function ParseDays(text, value)
    Dim i, character
    ParseDays = False
    If Len(text) = 0 Or Len(text) > 4 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseDays = (value >= 1 And value <= 3650)
End Function
