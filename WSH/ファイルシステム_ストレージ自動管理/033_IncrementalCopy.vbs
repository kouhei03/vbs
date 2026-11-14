' 概要: 更新日時またはサイズが異なるファイルだけをコピーします。宛先の余分なファイルは削除しません。
' 実行方法: Usage: cscript //nologo 033_IncrementalCopy.vbs <source-folder> <destination-folder> [APPLY]
' 処理結果: 更新日時・サイズに基づく差分コピー。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, source, destination, mode, file, target, changed
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In fso.GetFolder(source).Files
    target = fso.BuildPath(destination, file.Name)
    changed = Not fso.FileExists(target)
    If Not changed Then
        changed = (file.Size <> fso.GetFile(target).Size Or file.DateLastModified <> fso.GetFile(target).DateLastModified)
    End If
    If changed Then
        WScript.Echo file.Path & " -> " & target
        ' 主処理: 対象をコピーします。
        If mode = "APPLY" Then fso.CopyFile file.Path, target, True
    End If
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to copy."
