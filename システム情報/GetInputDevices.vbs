Option Explicit

Dim service, devices, device
Set service = GetObject("winmgmts:\\.\root\cimv2")

WScript.Echo "Keyboards:"
Set devices = service.ExecQuery("SELECT Description FROM Win32_Keyboard")
For Each device In devices
    WScript.Echo "  " & device.Description
Next

WScript.Echo "Pointing devices:"
Set devices = service.ExecQuery("SELECT Description FROM Win32_PointingDevice")
For Each device In devices
    WScript.Echo "  " & device.Description
Next
