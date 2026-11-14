' 概要: ロックファイルが消えるまで指定秒数以内で待機します（作成／削除は行いません）。
' 実行方法: Usage: cscript //nologo 039_WaitForLockFile.vbs <lock-file> <timeout-seconds-1..3600>
' 処理結果: ロックファイル消失の期限付き待機。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, lockPath, timeout, started
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
' 主処理: 条件を確認しながら処理を繰り返します。
Do While fso.FileExists(lockPath)
    If (Timer - started + 86400) Mod 86400 >= timeout Then
        WScript.Echo "Timed out waiting for lock removal: " & lockPath
        WScript.Quit 1
    End If
    WScript.Sleep 1000
Loop
WScript.Echo "Lock is absent: " & lockPath
