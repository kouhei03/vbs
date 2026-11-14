' 概要: 指定ファイルのサイズとSHA-256を表示します。FTP等への転送は行いません。
' 実行方法: Usage: cscript //nologo 024_LocalTransferManifest.vbs <file>
' 処理結果: 転送前のローカルファイルSHA-256表示。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, path, shell, runner, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 024_LocalTransferManifest.vbs <file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
WScript.Echo "Path=" & path & " Bytes=" & fso.GetFile(path).Size
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
Do While Not runner.StdOut.AtEndOfStream
    line = Trim(runner.StdOut.ReadLine)
    If line <> "" And InStr(1, line, "SHA256", 1) = 0 And InStr(1, line, "CertUtil:", 1) = 0 Then WScript.Echo "SHA256=" & line
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
