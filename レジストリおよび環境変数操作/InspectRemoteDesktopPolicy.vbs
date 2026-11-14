Option Explicit
Dim registry, result, denyConnections
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Control\Terminal Server", "fDenyTSConnections", denyConnections)
If result <> 0 Then
    WScript.Echo "RDP policy value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "fDenyTSConnections=" & denyConnections
