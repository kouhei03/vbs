' 概要: 指定した単一共有ファイルの更新日時・サイズ・ACL所有者情報を照会します。
' 実行方法: Usage: cscript //nologo 021_SharedFileMetadataAudit.vbs <file-path>
' 処理結果: 単一共有ファイルのメタデータとACL確認。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim path, re, fso, file, shell, runner, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 021_SharedFileMetadataAudit.vbs <file-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^(?:[A-Za-z]:\\|\\\\)[A-Za-z0-9 _.$-][A-Za-z0-9 _.$\\-]*$"
If Not re.Test(path) Then WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then WScript.Quit 1
Set file = fso.GetFile(path)
WScript.Echo "Path=" & file.Path & " Bytes=" & file.Size & " Modified=" & file.DateLastModified
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("icacls.exe """ & path & """")
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    WScript.Echo line
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
WScript.Echo "ACL output may show permissions, not the last user who modified the file."
