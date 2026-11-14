' 概要: ネットワーク共有上のファイルとローカル候補のサイズ・SHA-256を比較します。コピーはしません。
' 実行方法: Usage: cscript //nologo 038_NetworkFileCopyAudit.vbs <source-file> <destination-file>
' 処理結果: 2ファイルのサイズ・SHA-256比較（コピーなし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, source, destination, shell, sourceHash, destinationHash
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 038_NetworkFileCopyAudit.vbs <source-file> <destination-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(source) Or Not fso.FileExists(destination) Then WScript.Quit 2
WScript.Echo "SourceBytes=" & fso.GetFile(source).Size
WScript.Echo "DestinationBytes=" & fso.GetFile(destination).Size
sourceHash = HashFile(source)
destinationHash = HashFile(destination)
WScript.Echo "SourceSHA256=" & sourceHash
WScript.Echo "DestinationSHA256=" & destinationHash
If sourceHash <> destinationHash Then
    WScript.Echo "Mismatch: files differ."
    WScript.Quit 1
End If
WScript.Echo "Match."
Function HashFile(path)
    Dim runner, line, hash
    Set shell = CreateObject("WScript.Shell")
    ' 主処理: 外部コマンドの出力を読み取ります。
    Set runner = shell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
    hash = ""
    Do While Not runner.StdOut.AtEndOfStream
        line = Trim(runner.StdOut.ReadLine)
        If Len(line) = 64 And InStr(1, line, " ", 1) = 0 Then hash = line
    Loop
    If runner.ExitCode <> 0 Or hash = "" Then Err.Raise 5, , "Could not compute SHA-256 for " & path
    HashFile = hash
End Function
