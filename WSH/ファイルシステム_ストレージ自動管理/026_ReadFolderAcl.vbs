' icaclsで指定ファイル／フォルダーのACLを照会します（変更は行いません）。
Option Explicit
Dim fso, path, shell, process, line
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 026_ReadFolderAcl.vbs <file-or-folder>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) And Not fso.FolderExists(path) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("icacls.exe """ & path & """")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo process.StdErr.ReadAll
    WScript.Quit process.ExitCode
End If
