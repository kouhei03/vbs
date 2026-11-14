' 概要: 共有先へ新規／更新ファイルをコピーします。宛先の余分なファイルは削除しません。
' 実行方法: Usage: cscript //nologo 019_MirrorCopy.vbs <source-folder> <destination-folder> [APPLY]
' 処理結果: 更新ファイルの再帰コピー（宛先の余分なファイルは保持）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: APPLY指定時のみ実行し、更新時は既存宛先を上書きします。
Option Explicit
Dim fso, source, destination, mode
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 019_MirrorCopy.vbs <source-folder> <destination-folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = fso.GetAbsolutePathName(WScript.Arguments(1))
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If Not fso.FolderExists(source) Or IsSameOrChild(destination, source) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
MirrorFolder fso.GetFolder(source), destination
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to copy."
Sub MirrorFolder(sourceFolder, destinationPath)
    Dim file, child, target, copied
    If mode = "APPLY" And Not fso.FolderExists(destinationPath) Then fso.CreateFolder destinationPath
    ' 主処理: 対象の一覧を1件ずつ処理します。
    For Each file In sourceFolder.Files
        target = fso.BuildPath(destinationPath, file.Name)
        copied = Not fso.FileExists(target)
        If Not copied Then
            copied = file.Size <> fso.GetFile(target).Size Or file.DateLastModified > fso.GetFile(target).DateLastModified
        End If
        If copied Then
            WScript.Echo file.Path & " -> " & target
            ' 主処理: 対象をコピーします。
            If mode = "APPLY" Then fso.CopyFile file.Path, target, True
        End If
    Next
    For Each child In sourceFolder.SubFolders
        If (child.Attributes And 1024) = 0 Then MirrorFolder child, fso.BuildPath(destinationPath, child.Name)
    Next
End Sub
Function IsSameOrChild(path, parentPath)
    Dim normalizedPath, normalizedParent
    normalizedPath = LCase(path)
    normalizedParent = LCase(parentPath)
    If Right(normalizedPath, 1) = "\" And Len(normalizedPath) > 3 Then normalizedPath = Left(normalizedPath, Len(normalizedPath) - 1)
    If Right(normalizedParent, 1) = "\" And Len(normalizedParent) > 3 Then normalizedParent = Left(normalizedParent, Len(normalizedParent) - 1)
    If normalizedPath = normalizedParent Then
        IsSameOrChild = True
    ElseIf Right(normalizedParent, 1) = "\" Then
        IsSameOrChild = (Left(normalizedPath, Len(normalizedParent)) = normalizedParent)
    Else
        IsSameOrChild = (Left(normalizedPath, Len(normalizedParent) + 1) = normalizedParent & "\")
    End If
End Function
