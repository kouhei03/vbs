' 概要: certutilで指定ファイルのSHA-256を計算します（ファイルの変更なし）。
' 実行方法: Usage: cscript //nologo 009_FileHash.vbs <file>
' 処理結果: `certutil`によるSHA-256計算。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, path, shell, command, process, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 009_FileHash.vbs <file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
    WScript.Echo "File not found: " & path
    WScript.Quit 1
End If
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set process = shell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
Do Until process.StdOut.AtEndOfStream
    line = Trim(process.StdOut.ReadLine)
    If line <> "" And InStr(line, "SHA256") = 0 And InStr(line, "CertUtil:") = 0 Then WScript.Echo line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo process.StdErr.ReadAll
    WScript.Quit process.ExitCode
End If
