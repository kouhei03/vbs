' 一般的なブラウザキャッシュフォルダーの存在とサイズだけを確認します。削除しません。
Option Explicit
Dim fso, shell, appData, localAppData
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
appData = shell.ExpandEnvironmentStrings("%APPDATA%")
localAppData = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
ReportFolder fso.BuildPath(localAppData, "Google\Chrome\User Data\Default\Cache")
ReportFolder fso.BuildPath(localAppData, "Microsoft\Edge\User Data\Default\Cache")
ReportExists fso.BuildPath(appData, "Mozilla\Firefox\Profiles")
WScript.Echo "Audit only; browser data was not modified."
Sub ReportFolder(path)
    If fso.FolderExists(path) Then
        WScript.Echo path & vbTab & "Bytes=" & FolderBytes(fso.GetFolder(path))
    Else
        WScript.Echo path & vbTab & "(not found)"
    End If
End Sub
Sub ReportExists(path)
    If fso.FolderExists(path) Then
        WScript.Echo path & vbTab & "(profiles folder exists; size not traversed)"
    Else
        WScript.Echo path & vbTab & "(not found)"
    End If
End Sub
Function FolderBytes(folder)
    Dim file, child, total
    total = CDbl(0)
    For Each file In folder.Files
        total = total + CDbl(file.Size)
    Next
    For Each child In folder.SubFolders
        If (child.Attributes And 1024) = 0 Then total = total + FolderBytes(child)
    Next
    FolderBytes = total
End Function
