' 指定テキストを固定のNotepadアプリで開きます。任意コマンドは実行しません。
Option Explicit
Dim fso, path, shell
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 053_OpenTextFileInNotepad.vbs <text-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
    WScript.Echo "File not found: " & path
    WScript.Quit 1
End If
Set shell = CreateObject("WScript.Shell")
shell.Run "notepad.exe """ & path & """", 1, False
WScript.Echo "Opened in Notepad: " & path
