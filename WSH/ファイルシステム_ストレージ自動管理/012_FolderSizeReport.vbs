' 指定フォルダー以下のファイル数と合計サイズを表示します。
Option Explicit
Dim fso, root, fileCount, totalBytes, threshold
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 012_FolderSizeReport.vbs <folder> [threshold-bytes]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(root) Then WScript.Quit 2
threshold = -1
If WScript.Arguments.Count = 2 Then
    If Not IsNumeric(WScript.Arguments(1)) Then WScript.Quit 2
    threshold = CDbl(WScript.Arguments(1))
    If threshold < 0 Then WScript.Quit 2
End If
fileCount = 0
totalBytes = CDbl(0)
ScanFolder fso.GetFolder(root)
WScript.Echo "Files=" & fileCount & " Bytes=" & totalBytes
If threshold >= 0 And totalBytes > threshold Then
    WScript.Echo "WARNING: total size exceeds " & threshold & " bytes."
    WScript.Quit 1
End If
Sub ScanFolder(folder)
    Dim file, child
    For Each file In folder.Files
        fileCount = fileCount + 1
        totalBytes = totalBytes + CDbl(file.Size)
    Next
    For Each child In folder.SubFolders
        If (child.Attributes And 1024) = 0 Then ScanFolder child
    Next
End Sub
