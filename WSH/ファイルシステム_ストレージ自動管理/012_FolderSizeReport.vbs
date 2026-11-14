' 概要: 指定フォルダー以下のファイル数と合計サイズを表示します。
' 実行方法: Usage: cscript //nologo 012_FolderSizeReport.vbs <folder> [threshold-bytes]
' 処理結果: 再帰的なファイル数・合計サイズ集計と任意の閾値警告。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, root, fileCount, totalBytes, threshold
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
    ' 主処理: 対象の一覧を1件ずつ確認します。
    For Each file In folder.Files
        fileCount = fileCount + 1
        totalBytes = totalBytes + CDbl(file.Size)
    Next
    For Each child In folder.SubFolders
        If (child.Attributes And 1024) = 0 Then ScanFolder child
    Next
End Sub
