Option Explicit

Dim inputPath, fso, inputFile, expression, line, count, uniqueNames
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 168_ServerCollectionScopePreview.vbs <server-list-file>"
    WScript.Quit 2
End If
inputPath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(inputPath) Then
    WScript.Echo "Input file does not exist."
    WScript.Quit 1
End If
If fso.GetFile(inputPath).Size > 1048576 Then
    WScript.Echo "Input file exceeds the 1 MiB limit."
    WScript.Quit 2
End If
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
expression.IgnoreCase = True
Set uniqueNames = CreateObject("Scripting.Dictionary")
uniqueNames.CompareMode = vbTextCompare
Set inputFile = fso.OpenTextFile(inputPath, 1, False)
Do Until inputFile.AtEndOfStream
    line = Trim(inputFile.ReadLine)
    If Len(line) > 0 And Left(line, 1) <> "#" Then
        If expression.Test(line) Then uniqueNames(line) = True
    End If
Loop
inputFile.Close
count = uniqueNames.Count
WScript.Echo "Validated unique server names: " & count
WScript.Echo "No network access or remote log collection was performed."
