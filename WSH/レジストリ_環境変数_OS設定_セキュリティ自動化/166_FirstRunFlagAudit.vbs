' 固定のデモアプリキーに初回起動フラグがあるか照会します。値は削除しません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\FirstRunComplete"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Demo first-run flag is absent."
Else
    WScript.Echo "Demo first-run flag exists: " & value
End If
On Error GoTo 0
WScript.Echo "No application state was reset."
