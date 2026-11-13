Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT OSArchitecture FROM Win32_OperatingSystem")

For Each system In systems
    WScript.Echo "OS architecture: " & system.OSArchitecture
Next
