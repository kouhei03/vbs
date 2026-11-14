' デモ専用HKCU値が存在するか検証します。値の作成・修復はしません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\Demo\ExampleValue"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Demo value not present."
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "Demo value is present and readable."
