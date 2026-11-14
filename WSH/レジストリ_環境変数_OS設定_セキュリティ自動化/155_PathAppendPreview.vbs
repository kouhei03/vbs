' PATHへの追加をプレビューし、重複有無を判定します。PATHは変更しません。
Option Explicit
Dim candidate, shell, env, currentPath, parts, item, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 155_PathAppendPreview.vbs <existing-folder>"
    WScript.Quit 2
End If
candidate = WScript.Arguments(0)
If InStr(candidate, ";") > 0 Or InStr(candidate, """") > 0 Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("USER")
If Not CreateObject("Scripting.FileSystemObject").FolderExists(candidate) Then
    WScript.Echo "Folder not found: " & candidate
    WScript.Quit 2
End If
currentPath = env("PATH")
parts = Split(currentPath, ";")
found = False
For Each item In parts
    If StrComp(Trim(item), candidate, 1) = 0 Then found = True
Next
If found Then
    WScript.Echo "Already present; no change needed."
Else
    WScript.Echo "Preview: user PATH would append " & candidate
End If
WScript.Echo "No environment variable was modified."
