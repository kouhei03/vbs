Option Explicit

Dim registry, result, zoneValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000001, _
    "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", _
    "CurrentLevel", zoneValue)
If result = 0 Then
    WScript.Echo "Current Internet zone security level: " & zoneValue
ElseIf result = 2 Then
    WScript.Echo "Current Internet zone level is not explicitly set for this user."
Else
    WScript.Echo "Zone setting unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Read-only query; zone or trusted-site settings were not changed."
