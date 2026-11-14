' ファイル概要: 標準ハイブファイルのメタデータを確認する（ロード・修復なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectRegistryHiveFiles.vbs
' 実行結果: 標準ハイブファイルのメタデータを確認する（ロード・修復なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, shell, root, hiveName, filePath, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
root = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\Config"
' [処理解説] 標準ハイブのファイル情報を列挙します。ファイル情報だけではハイブの整合性を判定できません。
For Each hiveName In Array("SYSTEM", "SOFTWARE", "SAM", "SECURITY", "DEFAULT")
    filePath = root & "\" & hiveName
    If fso.FileExists(filePath) Then
        Set file = fso.GetFile(filePath)
        WScript.Echo filePath & " size=" & file.Size & " modified=" & file.DateLastModified
    Else
        WScript.Echo "Missing file: " & filePath
    End If
Next
WScript.Echo "File metadata cannot establish hive integrity; no repair was attempted."
