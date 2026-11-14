' 外部ツールを起動せず、テキストを読み込み→キーワード抽出→新規ファイル出力します。
Option Explicit
Dim fso, inputPath, outputPath, keyword, inputFile, outputFile, line, count
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 075_LocalTextPipeline.vbs <input.txt> <new-output.txt> <keyword>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
keyword = WScript.Arguments(2)
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Or keyword = "" Then WScript.Quit 2
Set inputFile = fso.OpenTextFile(inputPath, 1, False, -2)
Set outputFile = fso.CreateTextFile(outputPath, False, True)
count = 0
Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    If InStr(1, line, keyword, 1) > 0 Then outputFile.WriteLine line: count = count + 1
Loop
inputFile.Close
outputFile.Close
WScript.Echo "Pipeline output lines=" & count & " Path=" & outputPath
