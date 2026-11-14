' サンプル専用キーの存在を確認します。キー・値の削除は行いません。
Option Explicit
Dim shell, value
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
value = shell.RegRead("HKEY_CURRENT_USER\Software\WSH\SafeExamples\Marker")
If Err.Number = 0 Then
    WScript.Echo "Cleanup candidate: HKCU\Software\WSH\SafeExamples\Marker"
Else
    WScript.Echo "No sample cleanup candidate found."
End If
On Error GoTo 0
WScript.Echo "Preview only; registry data was not deleted."
