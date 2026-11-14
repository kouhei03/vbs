' 概要: 2つの既存ファイルのSHA-256を比較し、破損・不一致を検出します（ファイル変更なし）。
' 実行方法: Usage: cscript //nologo 049_ChecksumCompare.vbs <file1> <file2>
' 処理結果: 2ファイルのSHA-256比較。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim fso, firstPath, secondPath, firstHash, secondHash, shell
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 049_ChecksumCompare.vbs <file1> <file2>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
firstPath = fso.GetAbsolutePathName(WScript.Arguments(0))
secondPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(firstPath) Or Not fso.FileExists(secondPath) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
firstHash = HashFile(firstPath)
secondHash = HashFile(secondPath)
WScript.Echo "File1SHA256=" & firstHash
WScript.Echo "File2SHA256=" & secondHash
If LCase(firstHash) <> LCase(secondHash) Then
    WScript.Echo "Mismatch."
    WScript.Quit 1
End If
WScript.Echo "Match."
Function HashFile(path)
    Dim runner, line, result
    ' 主処理: 外部コマンドの出力を読み取ります。
    Set runner = shell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
    result = ""
    Do While Not runner.StdOut.AtEndOfStream
        line = Trim(runner.StdOut.ReadLine)
        If Len(line) = 64 And InStr(1, line, " ", 1) = 0 Then result = line
    Loop
    If runner.ExitCode <> 0 Or result = "" Then Err.Raise 5, , "SHA-256 failed for " & path
    HashFile = result
End Function
