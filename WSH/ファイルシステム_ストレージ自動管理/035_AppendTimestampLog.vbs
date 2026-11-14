' 日時付きで一行をログに追記します。改行はログ行に混入しないよう空白化します。
Option Explicit
Dim fso, path, message, file
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 035_AppendTimestampLog.vbs <log-file> <message>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
Set file = fso.OpenTextFile(path, 8, True, True)
file.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
    Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & vbTab & message
file.Close
WScript.Echo "Appended to " & path
