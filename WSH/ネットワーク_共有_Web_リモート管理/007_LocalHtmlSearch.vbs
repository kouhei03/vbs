' ローカルHTMLファイルを検索して一致行番号だけ表示します。Web取得はしません。
Option Explicit
Dim fso, path, keyword, file, line, lineNo, matches
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 007_LocalHtmlSearch.vbs <local-html-file> <keyword>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
keyword = WScript.Arguments(1)
If Not fso.FileExists(path) Or keyword = "" Then WScript.Quit 2
Set file = fso.OpenTextFile(path, 1, False, -2)
lineNo = 0
matches = 0
Do Until file.AtEndOfStream
    line = file.ReadLine
    lineNo = lineNo + 1
    If InStr(1, line, keyword, 1) > 0 Then
        WScript.Echo "Match at line " & lineNo
        matches = matches + 1
    End If
Loop
file.Close
WScript.Echo "Matches=" & matches
