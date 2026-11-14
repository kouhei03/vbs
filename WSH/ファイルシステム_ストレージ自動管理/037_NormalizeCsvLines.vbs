' 概要: CSVを行単位で読み直し、改行をCRLFに統一して新しいファイルへ出力します。
' 実行方法: Usage: cscript //nologo 037_NormalizeCsvLines.vbs <input.csv> <new-output.csv>
' 処理結果: CSVの末尾空白除去と改行統一。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
' 補足: 引用符内の改行を含むCSVは対象外です。入力ファイルは変更しません。
Option Explicit
Dim fso, inputPath, outputPath, inputFile, outputFile, line
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 037_NormalizeCsvLines.vbs <input.csv> <new-output.csv>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Then WScript.Quit 2
' 主処理: 入力ファイルを開きます。
Set inputFile = fso.OpenTextFile(inputPath, 1, False, -2)
Set outputFile = fso.CreateTextFile(outputPath, False, True)
Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    outputFile.WriteLine TrimRight(line)
Loop
inputFile.Close
outputFile.Close
WScript.Echo "Normalized line endings: " & outputPath
Function TrimRight(value)
    Dim re
    Set re = New RegExp
    re.Pattern = "[ \t]+$"
    TrimRight = re.Replace(value, "")
End Function
