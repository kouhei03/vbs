Option Explicit
Dim service, systems, system, registry, result, revision
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Caption, BuildNumber, Version FROM Win32_OperatingSystem")
For Each system In systems
    WScript.Echo system.Caption & " Build=" & system.BuildNumber & " Version=" & system.Version
Next
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR", revision)
If result = 0 Then
    WScript.Echo "UBR=" & revision
Else
    WScript.Echo "UBR unavailable; WMI code=" & result
End If
