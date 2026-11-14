' 空フォルダーを再帰的に検出します。削除はDELETEを指定した場合だけ行います。
' 指定したルート自体と再解析ポイントは削除しません。
Option Explicit
Dim fso, root, mode
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 042_RemoveEmptyFolders.vbs <folder> [DELETE]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
mode = ""
If WScript.Arguments.Count = 2 Then mode = UCase(WScript.Arguments(1))
If Not fso.FolderExists(root) Or (mode <> "" And mode <> "DELETE") Then WScript.Quit 2
ScanFolder fso.GetFolder(root)
If mode <> "DELETE" Then WScript.Echo "Preview only; append DELETE to remove empty folders."
Sub ScanFolder(folder)
    Dim children, child, path
    Set children = CreateObject("Scripting.Dictionary")
    For Each child In folder.SubFolders
        If (child.Attributes And 1024) = 0 Then children.Add child.Path, child.Name
    Next
    For Each path In children.Keys
        If fso.FolderExists(path) Then
            ScanFolder fso.GetFolder(path)
            If fso.FolderExists(path) Then
                Set child = fso.GetFolder(path)
                If child.Files.Count = 0 And child.SubFolders.Count = 0 Then
                    WScript.Echo path
                    If mode = "DELETE" Then fso.DeleteFolder path, True
                End If
            End If
        End If
    Next
End Sub
