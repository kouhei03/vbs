Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special FROM Win32_UserProfile")
WScript.Echo "SID" & vbTab & "Path" & vbTab & "Loaded" & vbTab & "Special"
For Each profile In profiles
    WScript.Echo profile.SID & vbTab & profile.LocalPath & vbTab & _
        profile.Loaded & vbTab & profile.Special
Next
