Option Explicit
Dim registry, result, value
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\DirectX", "Version", value)
If result = 0 Then
    WScript.Echo "Version=" & value
ElseIf result = 2 Then
    WScript.Echo "Version is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
