' 指定拡張子の現在のUserChoice ProgIdを照会します。関連付けは変更しません。
Option Explicit
Dim extension, re, shell, path, value
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 164_FileAssociationAudit.vbs <extension-without-dot>"
    WScript.Quit 2
End If
extension = LCase(WScript.Arguments(0))
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9]{1,10}$"
If Not re.Test(extension) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & extension & "\UserChoice\ProgId"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "UserChoice ProgId not found or protected by Windows."
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "." & extension & " -> " & value
