' ファイル概要: ロックファイル状態照会。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 161_LockFileStatus.vbs
' 実行結果: ロックファイル状態照会処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, shell, lockPath
Set fso=CreateObject("Scripting.FileSystemObject")
Set shell=CreateObject("WScript.Shell")
' [処理解説] TEMPにある固定名のロックファイルを確認します。ファイルの作成・削除はしません。
lockPath=shell.ExpandEnvironmentStrings("%TEMP%") & "\vbs-process-example.lock"
If fso.FileExists(lockPath) Then
    WScript.Echo "Lock file exists: " & lockPath
Else
    WScript.Echo "No lock file exists."
End If
