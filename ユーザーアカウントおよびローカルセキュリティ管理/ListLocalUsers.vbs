Option Explicit

Dim service, users, user
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")

WScript.Echo "Local user" & vbTab & "Disabled" & vbTab & _
    "Locked out" & vbTab & "Password expires"
For Each user In users
    WScript.Echo user.Name & vbTab & user.Disabled & vbTab & _
        user.Lockout & vbTab & user.PasswordExpires
Next
