' 実行中スクリプト自身の絶対パスと格納フォルダーを表示します。
Option Explicit
Dim fso, scriptPath
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetAbsolutePathName(WScript.ScriptFullName)
WScript.Echo "Script=" & scriptPath
WScript.Echo "Folder=" & fso.GetParentFolderName(scriptPath)
