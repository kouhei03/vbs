' BOMを優先判定してテキストをUTF-8へ変換します。BOMなしはsource-charset指定が必要です。
Option Explicit
Dim fso, inputPath, outputPath, fallbackCharset, binary, prefix, sourceCharset, textStream, output, content
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 024_ConvertTextEncoding.vbs <input> <new-output> [source-charset]"
    WScript.Echo "BOMなしの例: ... input.txt output.txt Shift_JIS"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
fallbackCharset = "Shift_JIS"
If WScript.Arguments.Count = 3 Then fallbackCharset = WScript.Arguments(2)
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Then WScript.Quit 2
Set binary = CreateObject("ADODB.Stream")
binary.Type = 1
binary.Open
binary.LoadFromFile inputPath
sourceCharset = fallbackCharset
If binary.Size >= 3 Then
    binary.Position = 0
    prefix = binary.Read(3)
    If AscB(MidB(prefix, 1, 1)) = 239 And AscB(MidB(prefix, 2, 1)) = 187 And AscB(MidB(prefix, 3, 1)) = 191 Then sourceCharset = "utf-8"
End If
If binary.Size >= 2 Then
    binary.Position = 0
    prefix = binary.Read(2)
    If AscB(MidB(prefix, 1, 1)) = 255 And AscB(MidB(prefix, 2, 1)) = 254 Then sourceCharset = "unicode"
End If
binary.Close
Set textStream = CreateObject("ADODB.Stream")
textStream.Type = 2
textStream.Charset = sourceCharset
textStream.Open
textStream.LoadFromFile inputPath
textStream.Position = 0
content = textStream.ReadText
textStream.Close
Set textStream = CreateObject("ADODB.Stream")
textStream.Type = 2
textStream.Charset = "utf-8"
textStream.Open
textStream.WriteText content
textStream.SaveToFile outputPath, 1
textStream.Close
WScript.Echo "Converted using " & sourceCharset & " -> utf-8: " & outputPath
