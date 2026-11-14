Option Explicit
Dim service, monitors, monitor, monitorCount
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set monitors = service.ExecQuery("SELECT Name, PNPDeviceID, ScreenWidth, ScreenHeight, Status FROM Win32_DesktopMonitor")
monitorCount = monitors.Count
WScript.Echo "WMI-reported monitor count: " & monitorCount
For Each monitor In monitors
    WScript.Echo SafeText(monitor.Name) & " " & SafeText(monitor.ScreenWidth) & "x" & SafeText(monitor.ScreenHeight)
    WScript.Echo "PNPDeviceID=" & SafeText(monitor.PNPDeviceID) & " Status=" & SafeText(monitor.Status)
Next
WScript.Echo "Monitor physical arrangement coordinates are not exposed by this query."

Function SafeText(value)
    If IsNull(value) Then
        SafeText = ""
    Else
        SafeText = CStr(value)
    End If
End Function
