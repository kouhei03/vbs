Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Name, Status, Availability, VideoModeDescription FROM Win32_VideoController")
For Each adapter In adapters
    WScript.Echo adapter.Name & " Status=" & adapter.Status & " Availability=" & adapter.Availability
    WScript.Echo "VideoModeDescription=" & adapter.VideoModeDescription
Next
WScript.Echo "WMI does not expose a definitive global hardware-acceleration on/off switch."
