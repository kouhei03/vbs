' 概要: インストーラー候補のファイル情報のみ表示します。インストール／サイレント実行はしません。
' 実行方法: Usage: cscript //nologo 064_InstallerFileAudit.vbs <installer.msi-or-exe>
' 処理結果: インストーラーファイル情報の確認（実行なし）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, file, version
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 064_InstallerFileAudit.vbs <installer.msi-or-exe>"
    WScript.Quit 2
End If
' 主処理: インストーラーファイル情報の確認（実行なし）。
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
