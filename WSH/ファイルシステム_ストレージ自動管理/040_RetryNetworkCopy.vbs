' 指定ファイルを指定先へ最大3回コピーします。既存宛先は上書きしません。
' APPLY指定時のみ実行し、宛先はユーザーが指定したUNCまたはローカルパスです。
Option Explicit
Dim fso, source, destination, tempPath, mode, attempt, copied, lastError
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    WScript.Echo "Usage: cscript //nologo 040_RetryNetworkCopy.vbs <source-file> <destination-file> [APPLY]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
source = fso.GetAbsolutePathName(WScript.Arguments(0))
destination = WScript.Arguments(1)
mode = ""
If WScript.Arguments.Count = 3 Then mode = UCase(WScript.Arguments(2))
If Not fso.FileExists(source) Or (mode <> "" And mode <> "APPLY") Then WScript.Quit 2
If fso.FileExists(destination) Then
    WScript.Echo "Destination exists; refusing to overwrite: " & destination
    WScript.Quit 1
End If
Do
    tempPath = destination & ".wsh-copy-" & fso.GetTempName
Loop While fso.FileExists(tempPath)
If mode <> "APPLY" Then
    WScript.Echo "Preview: " & source & " -> " & destination
    WScript.Echo "Append APPLY to copy."
    WScript.Quit 0
End If
copied = False
For attempt = 1 To 3
    On Error Resume Next
    fso.CopyFile source, tempPath, False
    If Err.Number = 0 Then
        If fso.GetFile(tempPath).Size = fso.GetFile(source).Size Then
            If Not fso.FileExists(destination) Then
                fso.MoveFile tempPath, destination
                copied = (Err.Number = 0)
            Else
                Err.Raise 58, , "Destination appeared during copy."
            End If
        Else
            Err.Raise 5, , "Copied file size does not match the source."
        End If
    End If
    If Err.Number <> 0 Then lastError = Err.Description
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    If copied Then
        On Error GoTo 0
        Exit For
    End If
    WScript.Echo "Copy attempt " & attempt & " failed: " & lastError
    Err.Clear
    On Error GoTo 0
    If attempt < 3 Then WScript.Sleep attempt * 2000
Next
If Not copied Then WScript.Quit 1
WScript.Echo "Copied: " & destination
