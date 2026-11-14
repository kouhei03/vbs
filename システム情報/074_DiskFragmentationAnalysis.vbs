' ファイル概要: ディスク断片化分析。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 074_DiskFragmentationAnalysis.vbs
' 実行結果: ディスク断片化分析処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, shell, drives, drive, process, line, found
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
found = False
For Each drive In fso.Drives
    If drive.DriveType = 2 And drive.IsReady Then
        found = True
        WScript.Echo "Analysis: " & drive.DriveLetter & ":"
        Set process = shell.Exec("defrag.exe " & drive.DriveLetter & ": /A /V")
        Do While Not process.StdOut.AtEndOfStream
            line = process.StdOut.ReadLine
            WScript.Echo line
        Loop
        Do While Not process.StdErr.AtEndOfStream
            line = process.StdErr.ReadLine
            WScript.Echo "ERROR: " & line
        Loop
        If process.ExitCode <> 0 Then WScript.Echo "defrag analysis exit code=" & process.ExitCode
    End If
Next
If Not found Then WScript.Echo "No ready fixed drives found."
