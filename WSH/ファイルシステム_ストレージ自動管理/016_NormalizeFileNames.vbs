' 概要: 全角ASCII文字と空白を半角化し、空白をアンダースコアへ置換します。
' 実行方法: Usage: cscript //nologo 016_NormalizeFileNames.vbs <folder> [APPLY]
' 処理結果: 全角ASCII・空白を半角化し、空白をアンダースコアへ置換。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, root, mode, folder, file, cleaned, target, collision
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 016_NormalizeFileNames.vbs <folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
mode = ""
If WScript.Arguments.Count = 2 Then mode = UCase(WScript.Arguments(1))
If Not fso.FolderExists(root) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
Set folder = fso.GetFolder(root)
collision = False
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In folder.Files
    cleaned = Replace(StrConv(file.Name, vbNarrow), ChrW(&H3000), "_")
    cleaned = Replace(Trim(cleaned), " ", "_")
    Do While InStr(cleaned, "__") > 0
        cleaned = Replace(cleaned, "__", "_")
    Loop
    target = fso.BuildPath(root, cleaned)
    If StrComp(file.Path, target, 1) <> 0 And fso.FileExists(target) Then collision = True
Next
If collision Then
    WScript.Echo "A normalized name would collide; no files were renamed."
    WScript.Quit 1
End If
For Each file In folder.Files
    cleaned = Replace(StrConv(file.Name, vbNarrow), ChrW(&H3000), "_")
    cleaned = Replace(Trim(cleaned), " ", "_")
    Do While InStr(cleaned, "__") > 0
        cleaned = Replace(cleaned, "__", "_")
    Loop
    If cleaned <> file.Name Then
        WScript.Echo file.Name & " -> " & cleaned
        If mode = "APPLY" Then file.Name = cleaned
    End If
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to rename."
