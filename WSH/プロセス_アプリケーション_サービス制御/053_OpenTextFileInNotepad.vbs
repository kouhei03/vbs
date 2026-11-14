' 概要: 指定テキストを固定のNotepadアプリで開きます。任意コマンドは実行しません。
' 実行方法: Usage: cscript //nologo 053_OpenTextFileInNotepad.vbs <text-file>
' 処理結果: 指定ファイルを固定のNotepadで開く。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, shell
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 053_OpenTextFileInNotepad.vbs <text-file>"
    WScript.Quit 2
End If
' 主処理: 指定ファイルを固定のNotepadで開く。
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
    WScript.Echo "File not found: " & path
    WScript.Quit 1
End If
Set shell = CreateObject("WScript.Shell")
shell.Run "notepad.exe """ & path & """", 1, False
WScript.Echo "Opened in Notepad: " & path
