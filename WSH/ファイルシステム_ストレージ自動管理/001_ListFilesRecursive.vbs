' 概要: 指定フォルダー以下のファイル一覧をUTF-16ログに出力します。
' 実行方法: Usage: cscript //nologo 001_ListFilesRecursive.vbs <folder> <log-file>
' 処理結果: 再帰ファイル一覧とサイズをログ出力。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, root, logPath, outFile
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 001_ListFilesRecursive.vbs <folder> <log-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
logPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FolderExists(root) Then
    WScript.Echo "Folder not found: " & root
    WScript.Quit 1
End If
If fso.FileExists(logPath) Then
    WScript.Echo "Refusing to overwrite: " & logPath
    WScript.Quit 2
End If
Set outFile = fso.CreateTextFile(logPath, False, True)
outFile.WriteLine "Path" & vbTab & "Bytes"
ListFolder fso.GetFolder(root), outFile, fso.GetFileName(logPath)
outFile.Close
WScript.Echo "Wrote " & logPath
Sub ListFolder(folder, output, excludedName)
    Dim file, child, attr
    ' 主処理: 対象の一覧を1件ずつ処理します。
    For Each file In folder.Files
        If LCase(file.Name) <> LCase(excludedName) Or LCase(folder.Path) <> LCase(fso.GetParentFolderName(logPath)) Then
            output.WriteLine file.Path & vbTab & file.Size
        End If
    Next
    For Each child In folder.SubFolders
        attr = child.Attributes
        If (attr And 1024) = 0 Then ListFolder child, output, excludedName
    Next
End Sub
