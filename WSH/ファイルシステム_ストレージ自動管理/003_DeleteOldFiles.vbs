' 指定日数より古いファイルの削除候補を表示します。DELETE指定時のみ削除します。
Option Explicit
Dim fso, root, days, mode, cutoff
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 003_DeleteOldFiles.vbs <folder> <days> [DELETE]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(root) Or Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Provide an existing folder and a numeric day count."
    WScript.Quit 2
End If
days = CLng(WScript.Arguments(1))
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If days < 1 Or days > 36500 Or (mode <> "" And mode <> "DELETE") Then
    WScript.Echo "Days must be 1..36500; optional mode must be DELETE."
    WScript.Quit 2
End If
cutoff = DateAdd("d", -days, Now)
ScanFolder fso.GetFolder(root)
If mode <> "DELETE" Then WScript.Echo "Preview only; append DELETE to remove listed files."
Sub ScanFolder(folder)
    Dim file, child
    For Each file In folder.Files
        If file.DateLastModified < cutoff Then
            WScript.Echo file.Path
            If mode = "DELETE" Then file.Delete True
        End If
    Next
    For Each child In folder.SubFolders
        If (child.Attributes And 1024) = 0 Then ScanFolder child
    Next
End Sub
