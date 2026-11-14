Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery( _
    "SELECT ServicePackMajorVersion, ServicePackMinorVersion FROM Win32_OperatingSystem")

For Each system In systems
    WScript.Echo "Service Pack: " & system.ServicePackMajorVersion & "." & _
        system.ServicePackMinorVersion
Next
