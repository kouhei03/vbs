' 指定アプリの現在のバージョンと署名照会対象パスを表示します。更新は開始しません。
Option Explicit
Dim fso, path, file
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 068_UpdatePreflightAudit.vbs <application-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set file = fso.GetFile(path)
WScript.Echo "Path=" & path
WScript.Echo "Version=" & fso.GetFileVersion(path)
WScript.Echo "LastModified=" & file.DateLastModified
WScript.Echo "Preflight only; no installer was launched."
