' ファイル概要: ミニダンプのメタデータを列挙する（移動・コピーなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 158_MiniDumpInventory.vbs
' 実行結果: ミニダンプのメタデータを列挙する（移動・コピーなし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim fso, shell, folderPath, folder, fileItem, count, totalBytes
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
folderPath = fso.BuildPath(shell.ExpandEnvironmentStrings("%SystemRoot%"), "Minidump")
count = 0
totalBytes = 0

' [処理解説] 標準フォルダーのDMPファイルを対象に、件数とサイズを集計します。
If fso.FolderExists(folderPath) Then
    Set folder = fso.GetFolder(folderPath)
    For Each fileItem In folder.Files
        If LCase(fso.GetExtensionName(fileItem.Name)) = "dmp" Then
            count = count + 1
            totalBytes = totalBytes + fileItem.Size
            WScript.Echo fileItem.Name & vbTab & fileItem.Size & vbTab & fileItem.DateLastModified
        End If
    Next
Else
    WScript.Echo "Standard minidump folder is unavailable."
End If
WScript.Echo "Minidump count: " & count & "; total bytes: " & totalBytes
WScript.Echo "Files were not moved or copied."
