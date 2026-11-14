' ファイルが他プロセスに占有されていないか、指定秒数以内で再試行します。
Option Explicit
Dim fso, path, timeout, started, file, locked, attempt
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 030_WaitForFileUnlock.vbs <file> <timeout-seconds-1..300>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Timeout must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If CDbl(WScript.Arguments(1)) < 1 Or CDbl(WScript.Arguments(1)) > 300 Then WScript.Quit 2
timeout = CLng(WScript.Arguments(1))
If Not fso.FileExists(path) Then WScript.Quit 2
started = Timer
Do
    locked = False
    On Error Resume Next
    Set file = fso.OpenTextFile(path, 8, False)
    If Err.Number <> 0 Then
        locked = True
        Err.Clear
    Else
        file.Close
    End If
    On Error GoTo 0
    If Not locked Then
        WScript.Echo "File is available: " & path
        WScript.Quit 0
    End If
    WScript.Sleep 1000
Loop While (Timer - started + 86400) Mod 86400 < timeout
WScript.Echo "Timed out waiting for file: " & path
WScript.Quit 1
