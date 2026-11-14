' 概要: 空フォルダーを再帰的に検出します。削除はDELETEを指定した場合だけ行います。
' 実行方法: Usage: cscript //nologo 042_RemoveEmptyFolders.vbs <folder> [DELETE]
' 処理結果: 空フォルダー検出／指定時の削除。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: 指定したルート自体と再解析ポイントは削除しません。
Option Explicit
Dim fso, root, mode
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
    ' 主処理: 対象の一覧を1件ずつ処理します。
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
                    ' 主処理: 条件に一致した対象を削除します。
                    If mode = "DELETE" Then fso.DeleteFolder path, True
                End If
            End If
        End If
    Next
End Sub
