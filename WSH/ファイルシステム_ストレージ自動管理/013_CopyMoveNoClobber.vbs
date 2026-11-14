' コピー／移動時に同名ファイルを避け、未使用の番号付き名前を選びます。
' 操作はAPPLY指定時だけ実行します。
Option Explicit
Dim fso, source, destination, action, mode, file, newPath, i, ext, stem
If WScript.Arguments.Count < 3 Or WScript.Arguments.Count > 4 Then
    WScript.Echo "Usage: cscript //nologo 013_CopyMoveNoClobber.vbs <file> <destination-folder> <copy|move> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = fso.GetAbsolutePathName(WScript.Arguments(1))
action = LCase(WScript.Arguments(2))
mode = ""
If WScript.Arguments.Count = 4 Then mode = UCase(WScript.Arguments(3))
If Not fso.FileExists(source) Or (action <> "copy" And action <> "move") Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
If Not fso.FolderExists(destination) Then
    If mode = "APPLY" Then fso.CreateFolder destination
    If mode <> "APPLY" Then WScript.Echo "Destination folder would be created: " & destination
End If
stem = fso.GetBaseName(source)
ext = fso.GetExtensionName(source)
newPath = fso.BuildPath(destination, fso.GetFileName(source))
i = 1
Do While fso.FileExists(newPath)
    If ext = "" Then
        newPath = fso.BuildPath(destination, stem & "_" & i)
    Else
        newPath = fso.BuildPath(destination, stem & "_" & i & "." & ext)
    End If
    i = i + 1
Loop
WScript.Echo action & ": " & source & " -> " & newPath
If mode = "APPLY" Then
    If action = "copy" Then fso.CopyFile source, newPath, False
    If action = "move" Then fso.MoveFile source, newPath
Else
    WScript.Echo "Preview only; append APPLY to proceed."
End If
