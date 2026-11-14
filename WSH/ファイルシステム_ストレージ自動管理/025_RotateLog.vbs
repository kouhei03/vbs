' 指定ログをタイムスタンプ付き名へ退避し、新しい空ログを作成します。
Option Explicit
Dim fso, path, mode, rotated, stamp, file, output
If WScript.Arguments.Count < 1 Or WScript.Arguments.Count > 2 Then
    WScript.Echo "Usage: cscript //nologo 025_RotateLog.vbs <log-file> [ROTATE]"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
mode = ""
If WScript.Arguments.Count = 2 Then mode = UCase(WScript.Arguments(1))
If Not fso.FileExists(path) Or (mode <> "" And mode <> "ROTATE") Then WScript.Quit 2
Set file = fso.GetFile(path)
stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
rotated = path & "." & stamp
WScript.Echo path & " -> " & rotated
If mode <> "ROTATE" Then
    WScript.Echo "Preview only; append ROTATE to proceed."
    WScript.Quit 0
End If
If fso.FileExists(rotated) Then
    WScript.Echo "Rotation target already exists."
    WScript.Quit 1
End If
file.Move rotated
Set output = fso.CreateTextFile(path, False, True)
output.Close
WScript.Echo "Created new log: " & path
