Option Explicit
Dim registry, result, enabled, server, autoConfig, displayValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", enabled)
If result = 0 Then WScript.Echo "ProxyEnable=" & enabled
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", server)
If result = 0 Then
    displayValue = server
    If InStr(displayValue, "@") > 0 Then displayValue = "(value omitted because it may contain credentials)"
    WScript.Echo "ProxyServer=" & displayValue
End If
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "AutoConfigURL", autoConfig)
If result = 0 Then WScript.Echo "AutoConfigURL=" & autoConfig
