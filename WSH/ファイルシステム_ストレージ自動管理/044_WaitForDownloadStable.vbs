' ダウンロードファイルのサイズが指定秒数安定するまで待機します（最大300秒）。
Option Explicit
Dim fso, path, stableSeconds, previousSize, stableSince, deadline, currentSize
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 044_WaitForDownloadStable.vbs <file> <stable-seconds-1..60>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Stable-seconds must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If CDbl(WScript.Arguments(1)) < 1 Or CDbl(WScript.Arguments(1)) > 60 Then WScript.Quit 2
stableSeconds = CLng(WScript.Arguments(1))
deadline = DateAdd("s", 300, Now)
previousSize = -1
stableSince = Now
Do While Now < deadline
    If fso.FileExists(path) Then
        currentSize = fso.GetFile(path).Size
        If currentSize <> previousSize Then
            previousSize = currentSize
            stableSince = Now
        ElseIf DateDiff("s", stableSince, Now) >= stableSeconds Then
            WScript.Echo "Stable size (" & currentSize & " bytes): " & path
            WScript.Quit 0
        End If
    End If
    WScript.Sleep 1000
Loop
WScript.Echo "Timed out waiting for a stable download: " & path
WScript.Quit 1
