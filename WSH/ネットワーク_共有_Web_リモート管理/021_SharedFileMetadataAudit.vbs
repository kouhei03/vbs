' 指定した単一共有ファイルの更新日時・サイズ・ACL所有者情報を照会します。
Option Explicit
Dim path, re, fso, file, shell, runner, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 021_SharedFileMetadataAudit.vbs <file-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^(?:[A-Za-z]:\\|\\\\)[A-Za-z0-9 _.$-][A-Za-z0-9 _.$\\-]*$"
If Not re.Test(path) Then WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then WScript.Quit 1
Set file = fso.GetFile(path)
WScript.Echo "Path=" & file.Path & " Bytes=" & file.Size & " Modified=" & file.DateLastModified
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("icacls.exe """ & path & """")
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    WScript.Echo line
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
WScript.Echo "ACL output may show permissions, not the last user who modified the file."
