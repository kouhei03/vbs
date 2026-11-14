Option Explicit
Dim registry, result, settingName, settingValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each settingName In Array("Hidden", "HideFileExt", "ShowSuperHidden")
    result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", settingName, settingValue)
    If result = 0 Then
        WScript.Echo settingName & "=" & settingValue
    ElseIf result = 2 Then
        WScript.Echo settingName & "=(not set)"
    Else
        WScript.Echo settingName & " read failed; WMI return code: " & result
    End If
Next
