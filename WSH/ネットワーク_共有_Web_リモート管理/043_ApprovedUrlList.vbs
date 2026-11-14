' 管理者が指定したURLを一覧表示します。ブラウザ起動や一斉アクセスはしません。
Option Explicit
Dim fso, path, file, line, re
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
