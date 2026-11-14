' 概要: icaclsで指定ファイル／フォルダーのACLを照会します（変更は行いません）。
' 実行方法: Usage: cscript //nologo 026_ReadFolderAcl.vbs <file-or-folder>
' 処理結果: `icacls`によるACL照会。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, path, shell, process, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 026_ReadFolderAcl.vbs <file-or-folder>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) And Not fso.FolderExists(path) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set process = shell.Exec("icacls.exe """ & path & """")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo process.StdErr.ReadAll
    WScript.Quit process.ExitCode
End If
