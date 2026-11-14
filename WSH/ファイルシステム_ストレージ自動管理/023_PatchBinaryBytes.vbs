' 単一ファイルの指定オフセットを16進バイト列で置換し、変更前の.bakを作成します。
Option Explicit
Dim fso, path, offsetText, hexText, backup, stream, bytes, i, byteCount, position
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 023_PatchBinaryBytes.vbs <file> <zero-based-offset> <hex-bytes>"
    WScript.Echo "Example: ... file.bin 16 A1B2FF"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
offsetText = WScript.Arguments(1)
hexText = UCase(Replace(WScript.Arguments(2), " ", ""))
If Not fso.FileExists(path) Or Not IsNumeric(offsetText) Or Len(hexText) = 0 Or Len(hexText) Mod 2 <> 0 Then WScript.Quit 2
For i = 1 To Len(hexText)
    If InStr("0123456789ABCDEF", Mid(hexText, i, 1)) = 0 Then
        WScript.Echo "Bytes must be hexadecimal."
        WScript.Quit 2
    End If
Next
position = CDbl(offsetText)
byteCount = Len(hexText) \ 2
Set stream = CreateObject("ADODB.Stream")
stream.Type = 1
stream.Open
stream.LoadFromFile path
If position < 0 Or position + byteCount > stream.Size Then
    stream.Close
    WScript.Echo "Patch range is outside the file."
    WScript.Quit 2
End If
backup = path & ".bak"
If fso.FileExists(backup) Then
    stream.Close
    WScript.Echo "Backup already exists; refusing to overwrite."
    WScript.Quit 1
End If
fso.CopyFile path, backup, False
ReDim bytes(byteCount - 1)
For i = 0 To byteCount - 1
    bytes(i) = CInt("&H" & Mid(hexText, i * 2 + 1, 2))
Next
stream.Position = position
stream.Write bytes
stream.SaveToFile path, 2
stream.Close
WScript.Echo "Patched " & byteCount & " byte(s); backup: " & backup
