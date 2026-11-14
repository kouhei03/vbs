' 拡張子別にファイルを移動します。APPLYを指定するまでプレビューのみです。
Option Explicit
Dim fso, sourceFolder, destination, mode, file, ext, target, newPath, i
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 002_SortByExtension.vbs <source-folder> <destination-folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
sourceFolder = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = fso.GetAbsolutePathName(WScript.Arguments(1))
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If Not fso.FolderExists(sourceFolder) Then
    WScript.Echo "Source folder not found."
    WScript.Quit 1
End If
If mode <> "" And mode <> "APPLY" Then
    WScript.Echo "The only accepted mode is APPLY."
    WScript.Quit 2
End If
If Not fso.FolderExists(destination) And mode = "APPLY" Then fso.CreateFolder destination
For Each file In fso.GetFolder(sourceFolder).Files
    ext = fso.GetExtensionName(file.Name)
    If ext = "" Then ext = "_no_extension"
    target = fso.BuildPath(destination, LCase(ext))
    If Not fso.FolderExists(target) Then
        If mode = "APPLY" Then fso.CreateFolder target
    End If
    newPath = fso.BuildPath(target, file.Name)
    i = 1
    Do While fso.FileExists(newPath)
        If fso.GetExtensionName(file.Name) = "" Then
            newPath = fso.BuildPath(target, fso.GetBaseName(file.Name) & "_" & i)
        Else
            newPath = fso.BuildPath(target, fso.GetBaseName(file.Name) & "_" & i & "." & fso.GetExtensionName(file.Name))
        End If
        i = i + 1
    Loop
    WScript.Echo file.Path & " -> " & newPath
    If mode = "APPLY" Then file.Move newPath
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to move files."
