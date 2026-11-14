' ロックファイルが消えるまで指定秒数以内で待機します（作成／削除は行いません）。
Option Explicit
Dim fso, lockPath, timeout, started
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 039_WaitForLockFile.vbs <lock-file> <timeout-seconds-1..3600>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Timeout must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
lockPath = fso.GetAbsolutePathName(WScript.Arguments(0))
If CDbl(WScript.Arguments(1)) < 1 Or CDbl(WScript.Arguments(1)) > 3600 Then WScript.Quit 2
timeout = CLng(WScript.Arguments(1))
started = Timer
Do While fso.FileExists(lockPath)
    If (Timer - started + 86400) Mod 86400 >= timeout Then
        WScript.Echo "Timed out waiting for lock removal: " & lockPath
        WScript.Quit 1
    End If
    WScript.Sleep 1000
Loop
WScript.Echo "Lock is absent: " & lockPath
