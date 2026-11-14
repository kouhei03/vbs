Option Explicit
Dim fso, file, filePath, firstLine, sectionCount, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectRegFile.vbs <reg-file>"
    WScript.Quit 2
End If
filePath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(filePath) Then
    WScript.Echo "File not found: " & filePath
    WScript.Quit 1
End If
Set file = fso.OpenTextFile(filePath, 1, False)
If file.AtEndOfStream Then
    file.Close
    WScript.Echo "The file is empty."
    WScript.Quit 1
End If
firstLine = Trim(file.ReadLine)
sectionCount = 0
Do Until file.AtEndOfStream
    line = Trim(file.ReadLine)
    If Left(line, 1) = "[" And Right(line, 1) = "]" Then sectionCount = sectionCount + 1
Loop
file.Close
WScript.Echo "Header: " & firstLine
WScript.Echo "Registry sections: " & sectionCount
WScript.Echo "This is a basic format check only; no registry changes were made."
