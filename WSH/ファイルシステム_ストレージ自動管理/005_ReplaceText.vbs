' テキストファイル内の文字列を置換し、元ファイルを.bakとして保存します。
Option Explicit
Dim fso, path, oldText, newText, inputFile, content, tempPath, backupPath, output
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 005_ReplaceText.vbs <file> <old-text> <new-text>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
oldText = WScript.Arguments(1)
newText = WScript.Arguments(2)
If Not fso.FileExists(path) Or oldText = "" Then
    WScript.Echo "File must exist and old-text must not be empty."
    WScript.Quit 2
End If
backupPath = path & ".bak"
tempPath = path & ".replace.tmp"
If fso.FileExists(backupPath) Or fso.FileExists(tempPath) Then
    WScript.Echo "Backup or temporary file already exists; refusing to overwrite."
    WScript.Quit 1
End If
Set inputFile = fso.OpenTextFile(path, 1, False, -2)
content = inputFile.ReadAll
inputFile.Close
If InStr(content, oldText) = 0 Then
    WScript.Echo "Text not found; file unchanged."
    WScript.Quit 1
End If
content = Replace(content, oldText, newText, 0, -1, 0)
Set output = fso.CreateTextFile(tempPath, False, True)
output.Write content
output.Close
fso.CopyFile path, backupPath, False
fso.DeleteFile path, True
fso.MoveFile tempPath, path
WScript.Echo "Updated; backup: " & backupPath
