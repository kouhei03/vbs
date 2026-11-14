' WindowsのTEMP配下にランダム名の一意な作業フォルダーを作成します。自動削除はしません。
Option Explicit
Dim fso, tempRoot, path, attempts
Set fso = CreateObject("Scripting.FileSystemObject")
tempRoot = fso.GetSpecialFolder(2).Path
For attempts = 1 To 10
    path = fso.BuildPath(tempRoot, "wsh_work_" & fso.GetTempName)
    If Not fso.FolderExists(path) And Not fso.FileExists(path) Then
        fso.CreateFolder path
        WScript.Echo path
        WScript.Quit 0
    End If
Next
WScript.Echo "Could not create a unique temporary folder."
WScript.Quit 1
