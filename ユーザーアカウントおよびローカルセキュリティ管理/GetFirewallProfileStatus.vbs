Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\standardcimv2")
Set profiles = service.ExecQuery( _
    "SELECT Name, Enabled FROM MSFT_NetFirewallProfile")

WScript.Echo "Firewall profile" & vbTab & "Enabled"
For Each profile In profiles
    WScript.Echo profile.Name & vbTab & profile.Enabled
Next
