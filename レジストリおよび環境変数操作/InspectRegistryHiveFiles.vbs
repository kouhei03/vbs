Option Explicit
Dim fso, shell, root, hiveName, filePath, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
root = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\Config"
For Each hiveName In Array("SYSTEM", "SOFTWARE", "SAM", "SECURITY", "DEFAULT")
    filePath = root & "\" & hiveName
    If fso.FileExists(filePath) Then
        Set file = fso.GetFile(filePath)
        WScript.Echo filePath & " size=" & file.Size & " modified=" & file.DateLastModified
    Else
        WScript.Echo "Missing file: " & filePath
    End If
Next
WScript.Echo "File metadata cannot establish hive integrity; no repair was attempted."
