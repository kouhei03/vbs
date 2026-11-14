' 概要: 保護リストに記載したファイルを除外して削除候補を表示します。削除操作はしません。
' 実行方法: Usage: cscript //nologo 048_ProtectWhitelistFiles.vbs <folder> <protected-names.txt>
' 処理結果: 保護対象を除外した削除候補表示（削除なし）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: リストは1行1ファイル名で、対象フォルダー直下のみを照合します。
Option Explicit
Dim fso, root, listPath, listFile, protected, line, file
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 048_ProtectWhitelistFiles.vbs <folder> <protected-names.txt>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
listPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FolderExists(root) Or Not fso.FileExists(listPath) Then WScript.Quit 2
Set protected = CreateObject("Scripting.Dictionary")
protected.CompareMode = 1
Set listFile = fso.OpenTextFile(listPath, 1, False, -2)
Do Until listFile.AtEndOfStream
    line = Trim(listFile.ReadLine)
    If line <> "" Then protected(line) = True
Loop
listFile.Close
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In fso.GetFolder(root).Files
    If protected.Exists(file.Name) Then
        WScript.Echo "PROTECTED " & file.Path
    Else
        WScript.Echo "CANDIDATE (not deleted) " & file.Path
    End If
Next
