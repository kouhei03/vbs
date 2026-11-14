' 概要: 指定されたフォルダーパスを、存在しない親階層を含めて作成します。
' 実行方法: Usage: cscript //nologo 010_CreateFolderTree.vbs <folder>
' 処理結果: 親階層を含むフォルダー作成。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 010_CreateFolderTree.vbs <folder>"
    WScript.Quit 2
End If
' 主処理: 親階層を含むフォルダー作成。
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If fso.FolderExists(path) Then
    WScript.Echo "Already exists: " & path
Else
    EnsureFolder path
    WScript.Echo "Created: " & path
End If
Sub EnsureFolder(folderPath)
    Dim parentPath
    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath = folderPath Or parentPath = "" Then Err.Raise 5, , "Cannot resolve parent folder."
    If Not fso.FolderExists(parentPath) Then EnsureFolder parentPath
    fso.CreateFolder folderPath
End Sub
