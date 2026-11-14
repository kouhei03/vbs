' 概要: 指定ファイルのバージョン情報を表示します。
' 実行方法: Usage: cscript //nologo 034_ReadFileVersion.vbs <file>
' 処理結果: ファイルバージョン情報表示。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 034_ReadFileVersion.vbs <file>"
    WScript.Quit 2
End If
' 主処理: ファイルバージョン情報表示。
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
WScript.Echo "File=" & path
WScript.Echo "FileVersion=" & fso.GetFileVersion(path)
