' このサンプルが使うHKCUのデモキーの存在だけを確認します。作成はしません。
Option Explicit
Dim shell, value
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
value = shell.RegRead("HKEY_CURRENT_USER\Software\WSH\SafeExamples\Marker")
If Err.Number = 0 Then
    WScript.Echo "Demo marker exists: " & value
Else
    WScript.Echo "Demo marker not found."
End If
On Error GoTo 0
WScript.Echo "Read-only check; no key or value was created."
