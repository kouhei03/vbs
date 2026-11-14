' SMTP送信の代わりにローカルログへ日時付きアラートを書き込みます。
Option Explicit
Dim fso, path, message, output
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 020_LocalAlertLog.vbs <log-file> <message>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
Set output = fso.OpenTextFile(path, 8, True, True)
output.WriteLine Now & vbTab & "ALERT" & vbTab & message
output.Close
WScript.Echo "Local alert recorded: " & path
