' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: ファイルシステム情報と既定ハイブファイル一覧を保持します。
Dim fso, shell, root, hiveName, filePath, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
root = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\Config"
' 解説: ハイブファイルの有無、サイズ、更新時刻を参照します。修復やロードはしません。
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
