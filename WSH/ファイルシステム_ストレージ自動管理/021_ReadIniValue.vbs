' 概要: INIのセクションとキーを検索して値を表示します（最初に一致したキー）。
' 実行方法: Usage: cscript //nologo 021_ReadIniValue.vbs <file.ini> <section> <key>
' 処理結果: INI形式のセクション・キー読み取り。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, sectionName, keyName, file, line, currentSection, pos
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 021_ReadIniValue.vbs <file.ini> <section> <key>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
sectionName = LCase(Trim(WScript.Arguments(1)))
keyName = LCase(Trim(WScript.Arguments(2)))
If Not fso.FileExists(path) Then WScript.Quit 2
' 主処理: 入力ファイルを開きます。
Set file = fso.OpenTextFile(path, 1, False, -2)
currentSection = ""
Do Until file.AtEndOfStream
    line = Trim(file.ReadLine)
    If Len(line) >= 2 And Left(line, 1) = "[" And Right(line, 1) = "]" Then
        currentSection = LCase(Trim(Mid(line, 2, Len(line) - 2)))
    ElseIf currentSection = sectionName And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
        pos = InStr(line, "=")
        If pos > 0 And LCase(Trim(Left(line, pos - 1))) = keyName Then
            WScript.Echo Trim(Mid(line, pos + 1))
            file.Close
            WScript.Quit 0
        End If
    End If
Loop
file.Close
WScript.Echo "Section/key not found."
WScript.Quit 1
