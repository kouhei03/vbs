' 概要: 指定フォルダー直下のファイルに属性を付与／解除します（サブフォルダーは対象外）。
' 実行方法: Usage: cscript //nologo 014_BulkFileAttributes.vbs <folder> <readonly|hidden|system|archive> <set|clear> [APPLY]
' 処理結果: 直下ファイル属性の付与／解除。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, root, attributeName, action, mode, folder, file, bit, updated
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 3 Or WScript.Arguments.Count > 4 Then
    WScript.Echo "Usage: cscript //nologo 014_BulkFileAttributes.vbs <folder> <readonly|hidden|system|archive> <set|clear> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
attributeName = LCase(WScript.Arguments(1))
action = LCase(WScript.Arguments(2))
mode = ""
If WScript.Arguments.Count = 4 Then mode = UCase(WScript.Arguments(3))
If Not fso.FolderExists(root) Or (action <> "set" And action <> "clear") Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
Select Case attributeName
    Case "readonly": bit = 1
    Case "hidden": bit = 2
    Case "system": bit = 4
    Case "archive": bit = 32
    Case Else: WScript.Echo "Unsupported attribute.": WScript.Quit 2
End Select
Set folder = fso.GetFolder(root)
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In folder.Files
    updated = file.Attributes
    If action = "set" Then
        updated = updated Or bit
    Else
        updated = updated And Not bit
    End If
    WScript.Echo file.Path & " Attributes=" & updated
    If mode = "APPLY" Then file.Attributes = updated
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to change attributes."
