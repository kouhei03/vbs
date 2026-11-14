' 大きなログを逐次読み込み、キーワードを含む行だけ新規ファイルへ抽出します。
Option Explicit
Dim fso, inputPath, outputPath, keyword, inputFile, outputFile, line, count
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 041_ExtractErrorLines.vbs <input.log> <new-output.log> <keyword>"
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
    If InStr(1, line, keyword, 1) > 0 Then
        outputFile.WriteLine line
        count = count + 1
    End If
Loop
inputFile.Close
outputFile.Close
WScript.Echo "Matched lines=" & count & " Output=" & outputPath
