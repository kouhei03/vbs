' 単純な1行1レコードCSV向け。引用符内のカンマ／改行を含むCSVは対象外です。
Option Explicit
Dim fso, inputPath, outputPath, columnNo, expected, inputFile, outputFile, line, fields, rowNo
If WScript.Arguments.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo 006_FilterCsvColumn.vbs <input.csv> <output.csv> <1-based-column> <value>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Or Not IsNumeric(WScript.Arguments(2)) Then WScript.Quit 2
columnNo = CLng(WScript.Arguments(2))
If columnNo < 1 Then WScript.Quit 2
expected = WScript.Arguments(3)
Set inputFile = fso.OpenTextFile(inputPath, 1, False, -2)
Set outputFile = fso.CreateTextFile(outputPath, False, True)
rowNo = 0
Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    fields = Split(line, ",")
    If rowNo = 0 Then
        If columnNo > UBound(fields) + 1 Then
            inputFile.Close
            outputFile.Close
            fso.DeleteFile outputPath, True
            WScript.Echo "Column is outside the header."
            WScript.Quit 2
        End If
        outputFile.WriteLine line
    ElseIf columnNo <= UBound(fields) + 1 Then
        If Trim(fields(columnNo - 1)) = expected Then outputFile.WriteLine line
    End If
    rowNo = rowNo + 1
Loop
inputFile.Close
outputFile.Close
WScript.Echo "Wrote " & outputPath
