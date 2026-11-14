' 概要: 実行中スクリプト自身の絶対パスと格納フォルダーを表示します。
' 実行方法: cscript //nologo 022_ScriptFolderPath.vbs
' 処理結果: 実行中スクリプトの絶対パス表示。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, scriptPath
' 主処理: 実行中スクリプトの絶対パス表示。
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetAbsolutePathName(WScript.ScriptFullName)
WScript.Echo "Script=" & scriptPath
WScript.Echo "Folder=" & fso.GetParentFolderName(scriptPath)
