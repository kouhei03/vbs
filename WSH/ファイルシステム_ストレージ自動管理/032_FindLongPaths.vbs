' 指定フォルダー以下を走査し、絶対パスが指定文字数を超える項目を表示します。
Option Explicit
Dim fso, root, limit
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 032_FindLongPaths.vbs <folder> <max-character-count>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Maximum path length must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
If CDbl(WScript.Arguments(1)) < 1 Or CDbl(WScript.Arguments(1)) > 32767 Then WScript.Quit 2
limit = CLng(WScript.Arguments(1))
If Not fso.FolderExists(root) Or limit < 1 Then WScript.Quit 2
ScanFolder fso.GetFolder(root)
Sub ScanFolder(folder)
    Dim file, child
    For Each file In folder.Files
        If Len(file.Path) > limit Then WScript.Echo Len(file.Path) & vbTab & file.Path
    Next
    For Each child In folder.SubFolders
        If Len(child.Path) > limit Then WScript.Echo Len(child.Path) & vbTab & child.Path
        If (child.Attributes And 1024) = 0 Then ScanFolder child
    Next
End Sub
