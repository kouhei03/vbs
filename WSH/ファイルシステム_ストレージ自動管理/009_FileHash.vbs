' certutilで指定ファイルのSHA-256を計算します（ファイルの変更なし）。
Option Explicit
Dim fso, path, shell, command, process, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 009_FileHash.vbs <file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then
    WScript.Echo "File not found: " & path
    WScript.Quit 1
End If
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("certutil.exe -hashfile """ & path & """ SHA256")
Do Until process.StdOut.AtEndOfStream
    line = Trim(process.StdOut.ReadLine)
    If line <> "" And InStr(line, "SHA256") = 0 And InStr(line, "CertUtil:") = 0 Then WScript.Echo line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo process.StdErr.ReadAll
    WScript.Quit process.ExitCode
End If
