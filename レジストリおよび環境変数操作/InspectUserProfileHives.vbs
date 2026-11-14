Option Explicit
Dim service, profiles, profile, fso, hivePath, hiveExists
Set service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")
Set profiles = service.ExecQuery("SELECT SID, LocalPath, Loaded FROM Win32_UserProfile WHERE Special = FALSE")
For Each profile In profiles
    hivePath = profile.LocalPath & "\NTUSER.DAT"
    hiveExists = fso.FileExists(hivePath)
    WScript.Echo "SID=" & profile.SID & " Loaded=" & profile.Loaded & " NTUSER.DAT exists=" & hiveExists
Next
WScript.Echo "No user hive was loaded, edited, or removed."
