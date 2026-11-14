Option Explicit
Dim registry, result, autoLogon, userName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", autoLogon)
If result <> 0 Then autoLogon = "(not set)"
WScript.Echo "AutoAdminLogon: " & autoLogon
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName", userName)
If result <> 0 Then userName = "(not set)"
WScript.Echo "DefaultUserName: " & userName
WScript.Echo "Password value was not read or displayed."
