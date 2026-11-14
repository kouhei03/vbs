' 指定フォルダー直下のファイル名に接頭辞／接尾辞を付けます。衝突があれば中止します。
Option Explicit
Dim fso, root, prefix, suffix, mode, folder, file, target, collision
If WScript.Arguments.Count < 3 Or WScript.Arguments.Count > 4 Then
    WScript.Echo "Usage: cscript //nologo 015_AddNameAffixes.vbs <folder> <prefix> <suffix> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
prefix = WScript.Arguments(1)
suffix = WScript.Arguments(2)
mode = ""
If WScript.Arguments.Count = 4 Then mode = UCase(WScript.Arguments(3))
If Not fso.FolderExists(root) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
Set folder = fso.GetFolder(root)
collision = False
For Each file In folder.Files
    target = fso.BuildPath(root, prefix & fso.GetBaseName(file.Name) & suffix)
    If fso.GetExtensionName(file.Name) <> "" Then target = target & "." & fso.GetExtensionName(file.Name)
    If fso.FileExists(target) Then collision = True
Next
If collision Then
    WScript.Echo "A target name already exists; no files were renamed."
    WScript.Quit 1
End If
For Each file In folder.Files
    target = fso.BuildPath(root, prefix & fso.GetBaseName(file.Name) & suffix)
    If fso.GetExtensionName(file.Name) <> "" Then target = target & "." & fso.GetExtensionName(file.Name)
    WScript.Echo file.Name & " -> " & fso.GetFileName(target)
    If mode = "APPLY" Then file.Name = fso.GetFileName(target)
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to rename."
