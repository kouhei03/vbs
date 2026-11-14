' 保護リストに記載したファイルを除外して削除候補を表示します。削除操作はしません。
' リストは1行1ファイル名で、対象フォルダー直下のみを照合します。
Option Explicit
Dim fso, root, listPath, listFile, protected, line, file
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 048_ProtectWhitelistFiles.vbs <folder> <protected-names.txt>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
listPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FolderExists(root) Or Not fso.FileExists(listPath) Then WScript.Quit 2
Set protected = CreateObject("Scripting.Dictionary")
protected.CompareMode = 1
Set listFile = fso.OpenTextFile(listPath, 1, False, -2)
Do Until listFile.AtEndOfStream
    line = Trim(listFile.ReadLine)
    If line <> "" Then protected(line) = True
Loop
listFile.Close
For Each file In fso.GetFolder(root).Files
    If protected.Exists(file.Name) Then
        WScript.Echo "PROTECTED " & file.Path
    Else
        WScript.Echo "CANDIDATE (not deleted) " & file.Path
    End If
Next
