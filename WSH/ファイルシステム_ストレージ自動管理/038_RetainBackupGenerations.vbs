' 概要: バックアップ世代を作成し、古い世代を指定数だけ残して削除します。
' 実行方法: Usage: cscript //nologo 038_RetainBackupGenerations.vbs <source-folder> <backup-parent> <keep-count> [APPLY]
' 処理結果: バックアップ世代作成と保持数管理。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: 既定はプレビュー。APPLY指定でコピーと削除を実行します。
Option Explicit
Dim fso, source, parent, keepCount, mode, stamp, backupRoot, folder, generations(), count, i, j, swap, cutoff, re
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 3 Or WScript.Arguments.Count > 4 Then
    WScript.Echo "Usage: cscript //nologo 038_RetainBackupGenerations.vbs <source-folder> <backup-parent> <keep-count> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
parent = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not IsNumeric(WScript.Arguments(2)) Then
    WScript.Echo "Keep-count must be numeric."
    WScript.Quit 2
End If
If CDbl(WScript.Arguments(2)) < 1 Or CDbl(WScript.Arguments(2)) > 100 Then
    WScript.Echo "Keep-count must be between 1 and 100."
    WScript.Quit 2
End If
keepCount = CLng(WScript.Arguments(2))
mode = ""
If WScript.Arguments.Count = 4 Then mode = UCase(WScript.Arguments(3))
If Not fso.FolderExists(source) Or keepCount < 1 Or keepCount > 100 Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
If IsSameOrChild(parent, source) Then
    WScript.Echo "Backup parent must not be the source folder or inside it."
    WScript.Quit 2
End If
Set re = New RegExp
re.Pattern = "^backup_[0-9]{8}_[0-9]{6}$"
stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
backupRoot = fso.BuildPath(parent, "backup_" & stamp)
If fso.FolderExists(backupRoot) Then
    WScript.Echo "Backup generation already exists: " & backupRoot
    WScript.Quit 1
End If
WScript.Echo "New generation: " & backupRoot
If mode = "APPLY" Then
    EnsureFolder parent
    fso.CopyFolder source, backupRoot, False
End If
count = 0
If fso.FolderExists(parent) Then
    ' 主処理: 対象の一覧を1件ずつ処理します。
    For Each folder In fso.GetFolder(parent).SubFolders
        If re.Test(folder.Name) And (folder.Attributes And 1024) = 0 Then
            ReDim Preserve generations(count)
            generations(count) = folder.Path
            count = count + 1
        End If
    Next
End If
For i = 0 To count - 2
    For j = i + 1 To count - 1
        If fso.GetFolder(generations(i)).DateCreated < fso.GetFolder(generations(j)).DateCreated Then
            swap = generations(i): generations(i) = generations(j): generations(j) = swap
        End If
    Next
Next
If mode = "APPLY" Then cutoff = keepCount Else cutoff = keepCount - 1
For i = cutoff To count - 1
    WScript.Echo "Remove old generation: " & generations(i)
    ' 主処理: 条件に一致した対象を削除します。
    If mode = "APPLY" Then fso.DeleteFolder generations(i), True
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to create and prune backups."
Sub EnsureFolder(folderPath)
    Dim parentPath
    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath = "" Or parentPath = folderPath Then Err.Raise 5, , "Cannot resolve backup parent."
    If Not fso.FolderExists(parentPath) Then EnsureFolder parentPath
    fso.CreateFolder folderPath
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
