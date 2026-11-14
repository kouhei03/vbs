' 概要: 指定フォルダー内の同一拡張子テキストを名前順に連結します。
' 実行方法: Usage: cscript //nologo 007_ConcatenateTextFiles.vbs <folder> <output-file> <extension>
' 処理結果: 同一拡張子のテキストを名前順に連結。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, root, outputPath, extension, folder, file, names(), count, i, j, swap, inputFile, outputFile
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 007_ConcatenateTextFiles.vbs <folder> <output-file> <extension>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
extension = LCase(Replace(WScript.Arguments(2), ".", ""))
If Not fso.FolderExists(root) Or extension = "" Or fso.FileExists(outputPath) Then WScript.Quit 2
count = 0
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In fso.GetFolder(root).Files
    If LCase(fso.GetExtensionName(file.Name)) = extension And LCase(file.Path) <> LCase(outputPath) Then
        ReDim Preserve names(count)
        names(count) = file.Path
        count = count + 1
    End If
Next
If count = 0 Then
    WScript.Echo "No matching input files."
    WScript.Quit 1
End If
For i = 0 To count - 2
    For j = i + 1 To count - 1
        If StrComp(fso.GetFileName(names(i)), fso.GetFileName(names(j)), 1) > 0 Then
            swap = names(i): names(i) = names(j): names(j) = swap
        End If
    Next
Next
Set outputFile = fso.CreateTextFile(outputPath, False, True)
For i = 0 To count - 1
    Set inputFile = fso.OpenTextFile(names(i), 1, False, -2)
    outputFile.WriteLine "===== " & fso.GetFileName(names(i)) & " ====="
    outputFile.Write inputFile.ReadAll
    outputFile.WriteLine
    inputFile.Close
Next
outputFile.Close
WScript.Echo "Concatenated " & count & " file(s) into " & outputPath
