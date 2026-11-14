' USBSTORサービスの起動値を照会します。ポリシー変更・USB履歴列挙はしません。
Option Explicit
Dim shell, value, path
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start"
On Error Resume Next
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "USBSTOR policy unavailable: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "USBSTOR Start=" & value & " (read-only)"
