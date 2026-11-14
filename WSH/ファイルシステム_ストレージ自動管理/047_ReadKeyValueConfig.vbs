' 概要: JSONライクな設定テキストのkey=value行を読み込みます。JSON文法の完全なパーサーではありません。
' 実行方法: Usage: cscript //nologo 047_ReadKeyValueConfig.vbs <config-file>
' 処理結果: `key=value`設定ファイルの読み取り。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, file, line, separator, key, value, settings
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 047_ReadKeyValueConfig.vbs <config-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set settings = CreateObject("Scripting.Dictionary")
settings.CompareMode = 1
Set file = fso.OpenTextFile(path, 1, False, -2)
Do Until file.AtEndOfStream
    line = Trim(file.ReadLine)
    If line <> "" And Left(line, 1) <> "#" And Left(line, 1) <> ";" Then
        separator = InStr(line, "=")
        If separator < 2 Then
            file.Close
            WScript.Echo "Invalid key=value line: " & line
            WScript.Quit 1
        End If
        key = Trim(Left(line, separator - 1))
        value = Trim(Mid(line, separator + 1))
        If settings.Exists(key) Then
            file.Close
            WScript.Echo "Duplicate key: " & key
            WScript.Quit 1
        End If
        settings.Add key, value
    End If
Loop
file.Close
' 主処理: 対象の一覧を1件ずつ処理します。
For Each key In settings.Keys
    WScript.Echo key & "=" & settings(key)
Next
