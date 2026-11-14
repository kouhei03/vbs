Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery( _
    "SELECT PartOfDomain, Domain, Workgroup FROM Win32_ComputerSystem")

For Each system In systems
    If system.PartOfDomain Then
        WScript.Echo "Join type: Domain"
        WScript.Echo "Domain: " & system.Domain
    Else
        WScript.Echo "Join type: Workgroup"
        WScript.Echo "Workgroup: " & system.Workgroup
    End If
Next
