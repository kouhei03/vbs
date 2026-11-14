' 概要: 複数CSVの先頭行（ヘッダー）がすべて一致するか確認します。
' 実行方法: Usage: cscript //nologo 046_CompareCsvHeaders.vbs <file1.csv> <file2.csv> [file3.csv ...]
' 処理結果: 複数CSVのヘッダー比較。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, baseline, i, path, file, header
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript //nologo 046_CompareCsvHeaders.vbs <file1.csv> <file2.csv> [file3.csv ...]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
baseline = ""
For i = 0 To WScript.Arguments.Count - 1
    path = fso.GetAbsolutePathName(WScript.Arguments(i))
    If Not fso.FileExists(path) Then
        WScript.Echo "File not found: " & path
        WScript.Quit 2
    End If
    ' 主処理: 入力ファイルを開いて内容を読み取ります。
    Set file = fso.OpenTextFile(path, 1, False, -2)
    If file.AtEndOfStream Then
        file.Close
        WScript.Echo "Empty file: " & path
        WScript.Quit 1
    End If
    header = file.ReadLine
    file.Close
    If i = 0 Then
        baseline = header
        WScript.Echo "Baseline: " & path
    ElseIf StrComp(header, baseline, 0) <> 0 Then
        WScript.Echo "MISMATCH: " & path
        WScript.Quit 1
    Else
        WScript.Echo "MATCH: " & path
    End If
Next
WScript.Echo "All headers match."
