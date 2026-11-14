' ファイル概要: ロックファイルの単一起動デモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 154_SingleInstanceLock.vbs
' 実行結果: ロックファイルの単一起動デモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, shell, lockPath, lockFile, errorNumber
Set fso=CreateObject("Scripting.FileSystemObject")
Set shell=CreateObject("WScript.Shell")
lockPath=shell.ExpandEnvironmentStrings("%TEMP%") & "\vbs-process-example.lock"
If fso.FileExists(lockPath) Then WScript.Echo "Lock exists: " & lockPath: WScript.Quit 1
On Error Resume Next
' [処理解説] 上書き引数をFalseにして、新しい出力ファイルのみを作成します。
Set lockFile=fso.CreateTextFile(lockPath,False,False)
errorNumber=Err.Number
Err.Clear
On Error GoTo 0
If errorNumber<>0 Then WScript.Echo "Could not create lock file; another instance may be running.": WScript.Quit 1
lockFile.WriteLine "Created by " & WScript.ScriptName & " at " & Now
lockFile.Close
WScript.Echo "Lock created: " & lockPath
WScript.Echo "This demonstration leaves the lock in place; remove it manually after verifying no instance is active."
