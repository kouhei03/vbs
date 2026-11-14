Option Explicit
Dim registry, result, value
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Control\SecureBoot\State", "UEFISecureBootEnabled", value)
If result = 0 Then
    WScript.Echo "UEFISecureBootEnabled=" & value
ElseIf result = 2 Then
    WScript.Echo "UEFISecureBootEnabled is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
