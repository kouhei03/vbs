' 概要: 大きなログを逐次読み込み、キーワードを含む行だけ新規ファイルへ抽出します。
' 実行方法: Usage: cscript //nologo 041_ExtractErrorLines.vbs <input.log> <new-output.log> <keyword>
' 処理結果: キーワードを含むログ行の抽出。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, inputPath, outputPath, keyword, inputFile, outputFile, line, count
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 041_ExtractErrorLines.vbs <input.log> <new-output.log> <keyword>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
keyword = WScript.Arguments(2)
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Or keyword = "" Then WScript.Quit 2
' 主処理: 入力ファイルを開きます。
Set inputFile = fso.OpenTextFile(inputPath, 1, False, -2)
Set outputFile = fso.CreateTextFile(outputPath, False, True)
count = 0
Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    If InStr(1, line, keyword, 1) > 0 Then
        outputFile.WriteLine line
        count = count + 1
    End If
Loop
inputFile.Close
outputFile.Close
WScript.Echo "Matched lines=" & count & " Output=" & outputPath
