' ファイル名中の{DATE}と{USER}を置換します。既存名との衝突がある場合は中止します。
Option Explicit
Dim fso, network, root, folder, file, newName, target, dateText, userName, collision
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
