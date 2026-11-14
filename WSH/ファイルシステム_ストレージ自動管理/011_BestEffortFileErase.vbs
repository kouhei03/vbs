' 概要: 単一ファイルのみを3回上書きして削除するベストエフォート例です。
' 実行方法: Usage: cscript //nologo 011_BestEffortFileErase.vbs <file> <CONFIRM:absolute-file-path>
' 処理結果: 単一ファイルのベストエフォート上書き・削除。処理結果や状態はコンソールに表示されます。
' コード解説: ADODB.Streamでデータを読み書きし、指定形式の結果を扱います。
' 補足: 注意: SSD/スナップショット等からの復元を保証する安全な消去はできません。
Option Explicit
Dim fso, path, confirm, file, size, stream, offset, chunkSize, n, buffer, i, passNo
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 011_BestEffortFileErase.vbs <file> <CONFIRM:absolute-file-path>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
confirm = WScript.Arguments(1)
If Not fso.FileExists(path) Or confirm <> "CONFIRM:" & path Then
    WScript.Echo "File missing or confirmation did not exactly match the absolute path."
    WScript.Quit 2
End If
Set file = fso.GetFile(path)
If (file.Attributes And 1) <> 0 Then file.Attributes = file.Attributes And Not 1
size = file.Size
If size > 268435456 Then
    WScript.Echo "Files larger than 256 MiB are not accepted by this in-memory example."
    WScript.Quit 2
End If
If size = 0 Then
    ' 主処理: 条件に一致した対象を削除します。
    file.Delete True
    WScript.Echo "Deleted empty file; secure erasure is not guaranteed."
    WScript.Quit 0
End If
Randomize
For passNo = 1 To 3
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    stream.Position = 0
    offset = 0
    Do While offset < size
        chunkSize = 65536
        If size - offset < chunkSize Then chunkSize = size - offset
        ReDim buffer(chunkSize - 1)
        For i = 0 To chunkSize - 1
            buffer(i) = Int(Rnd * 256)
        Next
        stream.Write buffer
        offset = offset + chunkSize
    Loop
    stream.SetEOS
    stream.SaveToFile path, 2
    stream.Close
Next
fso.DeleteFile path, True
WScript.Echo "Best-effort overwrite and delete completed; recovery may still be possible."
