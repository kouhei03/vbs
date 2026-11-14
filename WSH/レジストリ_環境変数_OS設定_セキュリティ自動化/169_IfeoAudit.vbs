' IFEOデバッガー設定を固定のnotepad.exeキーで監査します。新規設定は行いません。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe\Debugger"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "No debugger value reported for notepad.exe."
Else
    WScript.Echo "Debugger value exists for notepad.exe: " & value
End If
On Error GoTo 0
WScript.Echo "Read-only audit; no IFEO policy was created."
