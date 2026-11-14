Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special FROM Win32_UserProfile " & _
    "WHERE Special=False")

WScript.Echo "SID" & vbTab & "Profile path" & vbTab & "Loaded"
For Each profile In profiles
    WScript.Echo profile.SID & vbTab & profile.LocalPath & vbTab & profile.Loaded
Next
