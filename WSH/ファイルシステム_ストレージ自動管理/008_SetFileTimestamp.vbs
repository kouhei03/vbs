' 概要: ファイルの更新日時を設定します。対象ファイルの現在値を表示してから変更します。
' 実行方法: Usage: cscript //nologo 008_SetFileTimestamp.vbs <file> <new-date-time>
' 処理結果: ファイル更新日時の変更。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, timestamp, file
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 008_SetFileTimestamp.vbs <file> <new-date-time>"
    WScript.Echo "Example: 2026-11-14 09:30:00"
    WScript.Quit 2
End If
' 主処理: ファイル更新日時の変更。
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
