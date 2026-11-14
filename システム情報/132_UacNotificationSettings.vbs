Option Explicit
Dim registry, result, value, settingName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each settingName In Array("EnableLUA", "ConsentPromptBehaviorAdmin", "PromptOnSecureDesktop", "FilterAdministratorToken")
    result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", settingName, value)
    If result = 0 Then
        WScript.Echo settingName & "=" & value
    Else
        WScript.Echo settingName & " unavailable; WMI code=" & result
    End If
Next
