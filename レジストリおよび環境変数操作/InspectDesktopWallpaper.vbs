Option Explicit
Dim registry, result, wallpaper
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(&H80000001, "Control Panel\Desktop", "Wallpaper", wallpaper)
If result <> 0 Then
    WScript.Echo "Wallpaper value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Wallpaper=" & wallpaper
