Option Explicit
Dim registry, result, startValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Services\USBSTOR", "Start", startValue)
If result <> 0 Then
    WScript.Echo "USBSTOR Start value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "USBSTOR Start=" & startValue & " (read-only; no policy change made)"
