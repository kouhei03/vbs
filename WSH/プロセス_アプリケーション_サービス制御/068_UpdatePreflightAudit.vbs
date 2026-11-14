' 概要: 指定アプリの現在のバージョンと署名照会対象パスを表示します。更新は開始しません。
' 実行方法: Usage: cscript //nologo 068_UpdatePreflightAudit.vbs <application-file>
' 処理結果: バージョン・更新日時の事前確認（更新なし）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, file
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 068_UpdatePreflightAudit.vbs <application-file>"
    WScript.Quit 2
End If
' 主処理: バージョン・更新日時の事前確認（更新なし）。
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set file = fso.GetFile(path)
WScript.Echo "Path=" & path
WScript.Echo "Version=" & fso.GetFileVersion(path)
WScript.Echo "LastModified=" & file.DateLastModified
WScript.Echo "Preflight only; no installer was launched."
