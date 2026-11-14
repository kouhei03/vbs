' UACのEnableLUA値を照会します。通知レベルやUACは変更しません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "EnableLUA unavailable: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "EnableLUA=" & value & " (read-only)"
