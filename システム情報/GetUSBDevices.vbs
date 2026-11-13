Option Explicit

Dim service, devices, device, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_PnPEntity " & _
    "WHERE PNPClass='USB'")

found = False
For Each device In devices
    found = True
    WScript.Echo "Device: " & device.Name
    WScript.Echo "Manufacturer: " & device.Manufacturer
    WScript.Echo "Status: " & device.Status
    WScript.Echo ""
Next
If Not found Then WScript.Echo "No USB devices returned by WMI."
