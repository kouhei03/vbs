' 指定ファイルのバージョン情報を表示します。
Option Explicit
Dim fso, path
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 034_ReadFileVersion.vbs <file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
WScript.Echo "File=" & path
WScript.Echo "FileVersion=" & fso.GetFileVersion(path)
