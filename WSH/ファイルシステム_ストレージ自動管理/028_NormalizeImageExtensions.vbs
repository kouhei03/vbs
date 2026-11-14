' 概要: 画像ヘッダーを確認し、JPEG/PNG/GIF/BMPの拡張子を正規化します。
' 実行方法: Usage: cscript //nologo 028_NormalizeImageExtensions.vbs <folder> [APPLY]
' 処理結果: 画像ヘッダーに基づく拡張子正規化。処理結果や状態はコンソールに表示されます。
' コード解説: ADODB.Streamでデータを読み書きし、指定形式の結果を扱います。
Option Explicit
Dim fso, root, mode, folder, file, stream, data, kind, extension, target
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 028_NormalizeImageExtensions.vbs <folder> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetAbsolutePathName(WScript.Arguments(0))
mode = ""
If WScript.Arguments.Count = 2 Then mode = UCase(WScript.Arguments(1))
If Not fso.FolderExists(root) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
Set folder = fso.GetFolder(root)
' 主処理: 対象の一覧を1件ずつ処理します。
For Each file In folder.Files
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile file.Path
    If stream.Size >= 8 Then
        stream.Position = 0
        data = stream.Read(8)
        kind = ""
        If AscB(MidB(data, 1, 1)) = 255 And AscB(MidB(data, 2, 1)) = 216 Then kind = "jpg"
        If AscB(MidB(data, 1, 1)) = 137 And MidB(data, 2, 3) = "PNG" Then kind = "png"
        If MidB(data, 1, 3) = "GIF" Then kind = "gif"
        If AscB(MidB(data, 1, 1)) = 66 And AscB(MidB(data, 2, 1)) = 77 Then kind = "bmp"
        stream.Close
        If kind <> "" And LCase(fso.GetExtensionName(file.Name)) <> kind Then
            extension = fso.GetExtensionName(file.Name)
            target = fso.BuildPath(root, fso.GetBaseName(file.Name) & "." & kind)
            If Not fso.FileExists(target) Then
                WScript.Echo file.Name & " -> " & fso.GetFileName(target)
                If mode = "APPLY" Then file.Name = fso.GetFileName(target)
            Else
                WScript.Echo "Skipped collision: " & target
            End If
        End If
    Else
        stream.Close
    End If
Next
If mode <> "APPLY" Then WScript.Echo "Preview only; append APPLY to rename."
