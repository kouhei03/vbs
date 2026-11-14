' 指定したHKCU\Software配下の値を読み取ります。レジストリ変更はしません。
Option Explicit
Dim path, re, shell, value
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 151_RegistryValueAudit.vbs <HKCU\\Software\\...\\ValueName>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^HKEY_CURRENT_USER\\Software\\[A-Za-z0-9 _.-]+(\\[A-Za-z0-9 _.-]+)*$"
If Not re.Test(path) Then
    WScript.Echo "Only a value under HKCU\\Software is accepted."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Read failed or value does not exist: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo path & "=" & value
