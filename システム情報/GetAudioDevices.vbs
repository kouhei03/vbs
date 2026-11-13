Option Explicit

Dim service, devices, device
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_SoundDevice")

For Each device In devices
    WScript.Echo "Device: " & device.Name
    WScript.Echo "Manufacturer: " & device.Manufacturer
    WScript.Echo "Status: " & device.Status
    WScript.Echo ""
Next
