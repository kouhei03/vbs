Option Explicit
Dim fso, i, path, file
Set fso=CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count=0 Then WScript.Echo "Drop one or more files onto the script or pass file paths.": WScript.Quit 2
For i=0 To WScript.Arguments.Count-1
    path=WScript.Arguments(i)
    If fso.FileExists(path) Then
        Set file=fso.GetFile(path)
        WScript.Echo file.Path & " bytes=" & file.Size & " modified=" & file.DateLastModified
    Else
        WScript.Echo "Not a file: " & path
    End If
Next
