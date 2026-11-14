' インストーラー候補のファイル情報のみ表示します。インストール／サイレント実行はしません。
Option Explicit
Dim fso, path, file, version
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 064_InstallerFileAudit.vbs <installer.msi-or-exe>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set file = fso.GetFile(path)
If LCase(fso.GetExtensionName(path)) <> "msi" And LCase(fso.GetExtensionName(path)) <> "exe" Then WScript.Quit 2
version = fso.GetFileVersion(path)
WScript.Echo "Path=" & path
WScript.Echo "Bytes=" & file.Size
WScript.Echo "FileVersion=" & version
WScript.Echo "Audit only; the installer was not executed."
