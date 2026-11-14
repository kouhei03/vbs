' 概要: ファイル名中の{DATE}と{USER}を置換します。既存名との衝突がある場合は中止します。
' 実行方法: Usage: cscript //nologo 036_ExpandFilenamePlaceholders.vbs <folder>
' 処理結果: `{DATE}`・`{USER}`のファイル名展開。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, network, root, folder, file, newName, target, dateText, userName, collision
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 036_ExpandFilenamePlaceholders.vbs <folder>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
Set network = CreateObject("WScript.Network")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(root) Then WScript.Quit 2
dateText = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
userName = network.UserName
Set folder = fso.GetFolder(root)
collision = False
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In folder.Files
    newName = Replace(Replace(file.Name, "{DATE}", dateText), "{USER}", userName)
    target = fso.BuildPath(root, newName)
    If newName <> file.Name And fso.FileExists(target) Then collision = True
Next
If collision Then
    WScript.Echo "A target name already exists; no files were renamed."
    WScript.Quit 1
End If
For Each file In folder.Files
    newName = Replace(Replace(file.Name, "{DATE}", dateText), "{USER}", userName)
    If newName <> file.Name Then
        WScript.Echo file.Name & " -> " & newName
        file.Name = newName
    End If
Next
