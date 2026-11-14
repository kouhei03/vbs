' ファイルの更新日時を設定します。対象ファイルの現在値を表示してから変更します。
Option Explicit
Dim fso, path, timestamp, file
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 008_SetFileTimestamp.vbs <file> <new-date-time>"
    WScript.Echo "Example: 2026-11-14 09:30:00"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Or Not IsDate(WScript.Arguments(1)) Then
    WScript.Echo "Provide an existing file and a valid date/time."
    WScript.Quit 2
End If
timestamp = CDate(WScript.Arguments(1))
Set file = fso.GetFile(path)
WScript.Echo "Old: " & file.DateLastModified
WScript.Echo "New: " & timestamp
file.DateLastModified = timestamp
WScript.Echo "Updated: " & path
