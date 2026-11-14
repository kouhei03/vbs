' クラウド転送は行わず、ローカルファイルのSHA-256マニフェストを新規作成します。
Option Explicit
Dim fso, inputPath, outputPath, inputFile, output, path, shell, runner, line, hash
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 042_LocalTransferChecksum.vbs <paths-list.txt> <new-manifest.txt>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FileExists(inputPath) Or fso.FileExists(outputPath) Then WScript.Quit 2
Set inputFile = fso.OpenTextFile(inputPath, 1, False, -2)
Set output = fso.CreateTextFile(outputPath, False, True)
Set shell = CreateObject("WScript.Shell")
Do Until inputFile.AtEndOfStream
    path = Trim(inputFile.ReadLine)
    If path <> "" Then
        If Not fso.FileExists(path) Then
            inputFile.Close
            output.Close
            fso.DeleteFile outputPath, True
            WScript.Echo "File not found: " & path
            WScript.Quit 1
        End If
        Set runner = shell.Exec("certutil.exe -hashfile """ & fso.GetAbsolutePathName(path) & """ SHA256")
        hash = ""
        Do While Not runner.StdOut.AtEndOfStream
            line = Trim(runner.StdOut.ReadLine)
            If Len(line) = 64 And InStr(1, line, " ", 1) = 0 Then hash = line
        Loop
        If runner.ExitCode <> 0 Or hash = "" Then
            inputFile.Close
            output.Close
            fso.DeleteFile outputPath, True
            WScript.Echo "Could not compute hash for: " & path
            WScript.Quit 1
        End If
        output.WriteLine hash & vbTab & fso.GetAbsolutePathName(path) & vbTab & fso.GetFile(path).Size
    End If
Loop
inputFile.Close
output.Close
WScript.Echo "Local-only manifest created: " & outputPath
