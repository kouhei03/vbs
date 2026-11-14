Option Explicit

Dim service, bioses, bios
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set bioses = service.ExecQuery( _
    "SELECT Manufacturer, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS")

For Each bios In bioses
    WScript.Echo "Manufacturer: " & bios.Manufacturer
    WScript.Echo "Version: " & bios.SMBIOSBIOSVersion
    WScript.Echo "Release date (WMI format): " & bios.ReleaseDate
Next
