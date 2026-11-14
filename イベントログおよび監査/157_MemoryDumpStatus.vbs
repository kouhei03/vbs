' ファイル概要: 標準のMEMORY.DMPのメタデータを確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 157_MemoryDumpStatus.vbs
' 実行結果: 標準のMEMORY.DMPのメタデータを確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shell, fso, systemRoot, dumpPath, dumpFile
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
systemRoot = shell.ExpandEnvironmentStrings("%SystemRoot%")
dumpPath = fso.BuildPath(systemRoot, "MEMORY.DMP")

If fso.FileExists(dumpPath) Then
    Set dumpFile = fso.GetFile(dumpPath)
    WScript.Echo "MEMORY.DMP exists; bytes=" & dumpFile.Size & _
        "; modified=" & dumpFile.DateLastModified
Else
    WScript.Echo "MEMORY.DMP was not found at the standard path."
End If
WScript.Echo "The dump was not opened, copied, or uploaded."
