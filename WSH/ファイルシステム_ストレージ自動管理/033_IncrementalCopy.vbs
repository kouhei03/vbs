' 更新日時またはサイズが異なるファイルだけをコピーします。宛先の余分なファイルは削除しません。
Option Explicit
Dim fso, source, destination, mode, file, target, changed
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 033_IncrementalCopy.vbs <source-folder> <destination-folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = fso.GetAbsolutePathName(WScript.Arguments(1))
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If Not fso.FolderExists(source) Or LCase(source) = LCase(destination) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
If Not fso.FolderExists(destination) And mode = "APPLY" Then fso.CreateFolder destination
For Each file In fso.GetFolder(source).Files
    target = fso.BuildPath(destination, file.Name)
    changed = Not fso.FileExists(target)
    If Not changed Then
        changed = (file.Size <> fso.GetFile(target).Size Or file.DateLastModified <> fso.GetFile(target).DateLastModified)
    End If
    If changed Then
        WScript.Echo file.Path & " -> " & target
        If mode = "APPLY" Then fso.CopyFile file.Path, target, True
    End If
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to copy."
