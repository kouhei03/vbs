' 指定フォルダー直下のファイルを名前順に連番化します。衝突確認後、APPLYで実行します。
Option Explicit
Dim fso, root, mode, folder, file, names(), tempNames(), tempSet, count, i, j, swap, target, stem, width, tempPath
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 017_NumberFiles.vbs <folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
mode = ""
If WScript.Arguments.Count = 2 Then mode = UCase(WScript.Arguments(1))
If Not fso.FolderExists(root) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
count = 0
For Each file In fso.GetFolder(root).Files
    ReDim Preserve names(count)
    names(count) = file.Path
    count = count + 1
Next
If count = 0 Then
    WScript.Echo "No files to number."
    WScript.Quit 0
End If
For i = 0 To count - 2
    For j = i + 1 To count - 1
        If StrComp(fso.GetFileName(names(i)), fso.GetFileName(names(j)), 1) > 0 Then
            swap = names(i): names(i) = names(j): names(j) = swap
        End If
    Next
Next
width = Len(CStr(count))
For i = 0 To count - 1
    Set file = fso.GetFile(names(i))
    stem = Right(String(width, "0") & CStr(i + 1), width) & "_" & file.Name
    target = fso.BuildPath(root, stem)
    If fso.FileExists(target) Then
        WScript.Echo "Name collision: " & target
        WScript.Quit 1
    End If
    WScript.Echo file.Name & " -> " & stem
Next
If mode <> "APPLY" Then
    WScript.Echo "Preview only; append APPLY to rename."
    WScript.Quit 0
End If
' 一時名を経由して、元ファイル名との衝突を避けます。
ReDim tempNames(count - 1)
Set tempSet = CreateObject("Scripting.Dictionary")
tempSet.CompareMode = 1
For i = 0 To count - 1
    Do
        tempNames(i) = fso.GetTempName
        tempPath = fso.BuildPath(root, tempNames(i))
    Loop While fso.FileExists(tempPath) Or tempSet.Exists(tempNames(i))
    tempSet.Add tempNames(i), True
Next
For i = 0 To count - 1
    Set file = fso.GetFile(names(i))
    file.Name = tempNames(i)
Next
For i = 0 To count - 1
    Set file = fso.GetFile(fso.BuildPath(root, tempNames(i)))
    file.Name = Right(String(width, "0") & CStr(i + 1), width) & "_" & fso.GetFileName(names(i))
Next
