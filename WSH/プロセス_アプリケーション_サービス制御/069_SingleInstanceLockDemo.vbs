' TEMP内の排他ファイルで単一起動を短時間だけデモします。
' 異常終了時に残ったロックは、稼働中インスタンスがないことを確認して手動削除してください。
Option Explicit
Dim fso, lockPath, lockFile, errorNumber
Set fso = CreateObject("Scripting.FileSystemObject")
lockPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, "wsh-069-single-instance.lock")
If fso.FileExists(lockPath) Then
    WScript.Echo "Another instance or stale lock exists: " & lockPath
    WScript.Quit 1
End If
On Error Resume Next
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
