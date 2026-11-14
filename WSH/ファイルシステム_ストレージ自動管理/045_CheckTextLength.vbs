' テキストの文字数を調べ、最小文字数未満ならエラー終了します。
Option Explicit
Dim fso, path, minimum, file, content
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 045_CheckTextLength.vbs <text-file> <minimum-characters>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Minimum character count must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Or CDbl(WScript.Arguments(1)) < 0 Or CDbl(WScript.Arguments(1)) > 2147483647 Then WScript.Quit 2
minimum = CLng(WScript.Arguments(1))
Set file = fso.OpenTextFile(path, 1, False, -2)
content = file.ReadAll
file.Close
WScript.Echo "Characters=" & Len(content) & " Minimum=" & minimum
If Len(content) < minimum Then
    WScript.Echo "ERROR: text is shorter than required."
    WScript.Quit 1
End If
