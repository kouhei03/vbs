' 概要: 管理者が指定したURLを一覧表示します。ブラウザ起動や一斉アクセスはしません。
' 実行方法: Usage: cscript //nologo 043_ApprovedUrlList.vbs <urls.txt>
' 処理結果: URLリストの検証・表示（ブラウザー起動なし）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, file, line, re
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 043_ApprovedUrlList.vbs <urls.txt>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set re = New RegExp
re.Pattern = "^https?://[^ ]+$"
re.IgnoreCase = True
' 主処理: 入力ファイルを開きます。
Set file = fso.OpenTextFile(path, 1, False, -2)
Do Until file.AtEndOfStream
    line = Trim(file.ReadLine)
    If line <> "" Then
        If Not re.Test(line) Then
            file.Close
            WScript.Echo "Invalid URL entry."
            WScript.Quit 1
        End If
        WScript.Echo line
    End If
Loop
file.Close
WScript.Echo "URLs listed only; none were opened."
