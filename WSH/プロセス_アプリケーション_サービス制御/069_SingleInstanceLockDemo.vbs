' 概要: TEMP内の排他ファイルで単一起動を短時間だけデモします。
' 実行方法: cscript //nologo 069_SingleInstanceLockDemo.vbs
' 処理結果: 一時ロックファイルによる単一起動デモ。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: 異常終了時に残ったロックは、稼働中インスタンスがないことを確認して手動削除してください。
Option Explicit
Dim fso, lockPath, lockFile, errorNumber
Set fso = CreateObject("Scripting.FileSystemObject")
lockPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, "wsh-069-single-instance.lock")
If fso.FileExists(lockPath) Then
    WScript.Echo "Another instance or stale lock exists: " & lockPath
    WScript.Quit 1
End If
On Error Resume Next
' 主処理: 結果ファイルを作成します。
Set lockFile = fso.CreateTextFile(lockPath, False, True)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
If errorNumber <> 0 Then WScript.Echo "Could not acquire lock.": WScript.Quit 1
lockFile.WriteLine "Created=" & Now & " Script=" & WScript.ScriptName
lockFile.Close
WScript.Echo "Lock acquired for a 2-second demonstration."
WScript.Sleep 2000
If fso.FileExists(lockPath) Then fso.DeleteFile lockPath, True
WScript.Echo "Lock released."
