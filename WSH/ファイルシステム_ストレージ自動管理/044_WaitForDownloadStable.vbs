' 概要: ダウンロードファイルのサイズが指定秒数安定するまで待機します（最大300秒）。
' 実行方法: Usage: cscript //nologo 044_WaitForDownloadStable.vbs <file> <stable-seconds-1..60>
' 処理結果: ダウンロードファイルのサイズ安定待機。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, stableSeconds, previousSize, stableSince, deadline, currentSize
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
' 主処理: 条件を確認しながら処理を繰り返します。
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
