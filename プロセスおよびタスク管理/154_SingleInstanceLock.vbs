' 解説: TEMPに排他ロックファイルを作成して単一起動を示します。既存ファイルは上書きしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, shell, lockPath, lockFile, errorNumber
Dim fso, shell, lockPath, lockFile, errorNumber
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso=CreateObject("Scripting.FileSystemObject")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
lockPath=shell.ExpandEnvironmentStrings("%TEMP%") & "\vbs-process-example.lock"
' 解説: 条件を判定し、必要な分岐を選択します: If fso.FileExists(lockPath)
If fso.FileExists(lockPath) Then WScript.Echo "Lock exists: " & lockPath: WScript.Quit 1
' 解説: この文でローカル診断処理を進めます: On Error Resume Next
On Error Resume Next
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set lockFile=fso.CreateTextFile(lockPath,False,False)
' 解説: 取得値または計算結果を変数に格納します。
errorNumber=Err.Number
' 解説: この文でローカル診断処理を進めます: Err.Clear
Err.Clear
' 解説: この文でローカル診断処理を進めます: On Error GoTo 0
On Error GoTo 0
' 解説: 条件を判定し、必要な分岐を選択します: If errorNumber<>0
If errorNumber<>0 Then WScript.Echo "Could not create lock file; another instance may be running.": WScript.Quit 1
' 解説: この文でローカル診断処理を進めます: lockFile.WriteLine "Created by " & WScript.ScriptName & " at " & Now
lockFile.WriteLine "Created by " & WScript.ScriptName & " at " & Now
' 解説: この文でローカル診断処理を進めます: lockFile.Close
lockFile.Close
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Lock created: " & lockPath
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "This demonstration leaves the lock in place; remove it manually after verifying no instance is active."
