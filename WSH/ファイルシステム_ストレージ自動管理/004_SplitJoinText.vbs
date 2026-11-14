' テキストを行単位で分割／結合します。改行形式は出力時にCRLFへ統一されます。
Option Explicit
Dim fso, mode, source, value, stream, lineNo, partNo, output, partPath, partName, outputPath, hasLaterPart
If WScript.Arguments.Count < 3 Or WScript.Arguments.Count > 4 Then
    WScript.Echo "Split: cscript //nologo 004_SplitJoinText.vbs split <file> <lines-per-part>"
    WScript.Echo "Join:  cscript //nologo 004_SplitJoinText.vbs join <parts-folder> <original-name> <output-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
mode = LCase(WScript.Arguments(0))
source = fso.GetAbsolutePathName(WScript.Arguments(1))
value = WScript.Arguments(2)
If mode = "split" Then
    If WScript.Arguments.Count <> 3 Then WScript.Quit 2
    If Not fso.FileExists(source) Or Not IsNumeric(value) Then WScript.Quit 2
    If CLng(value) < 1 Or CLng(value) > 1000000 Then WScript.Quit 2
    Set stream = fso.OpenTextFile(source, 1, False, -2)
    lineNo = 0
    partNo = 1
    Set output = Nothing
    Do Until stream.AtEndOfStream
        If lineNo Mod CLng(value) = 0 Then
            If Not output Is Nothing Then output.Close
            partPath = source & "." & Right("000" & partNo, 3) & ".part"
            If fso.FileExists(partPath) Then
                stream.Close
                WScript.Echo "Refusing to overwrite: " & partPath
                WScript.Quit 1
            End If
            Set output = fso.CreateTextFile(partPath, False, True)
            partNo = partNo + 1
        End If
        output.WriteLine stream.ReadLine
        lineNo = lineNo + 1
    Loop
    If Not output Is Nothing Then output.Close
    stream.Close
ElseIf mode = "join" Then
    If WScript.Arguments.Count <> 4 Then WScript.Quit 2
    If Not fso.FolderExists(source) Then WScript.Quit 2
    If value = "" Or fso.GetFileName(value) <> value Then
        WScript.Echo "Original-name must be a file name, not a path."
        WScript.Quit 2
    End If
    outputPath = fso.GetAbsolutePathName(WScript.Arguments(3))
    If fso.FileExists(outputPath) Then
        WScript.Echo "Refusing to overwrite: " & outputPath
        WScript.Quit 2
    End If
    Set output = fso.CreateTextFile(outputPath, False, True)
    For partNo = 1 To 999
        partName = fso.BuildPath(source, value & "." & Right("000" & partNo, 3) & ".part")
        If Not fso.FileExists(partName) Then
            hasLaterPart = False
            For lineNo = partNo + 1 To 999
                If fso.FileExists(fso.BuildPath(source, value & "." & Right("000" & lineNo, 3) & ".part")) Then hasLaterPart = True
            Next
            If hasLaterPart Then
                output.Close
                fso.DeleteFile outputPath, True
                WScript.Echo "Part sequence has a gap at " & partNo
                WScript.Quit 1
            End If
            Exit For
        End If
        Set stream = fso.OpenTextFile(partName, 1, False, -2)
        Do Until stream.AtEndOfStream
            output.WriteLine stream.ReadLine
        Loop
        stream.Close
    Next
    output.Close
    If partNo = 1 Then
        fso.DeleteFile outputPath, True
        WScript.Echo "No matching .part files found."
        WScript.Quit 1
    End If
Else
    WScript.Echo "Mode must be split or join."
    WScript.Quit 2
End If
