Option Explicit

Dim service, users, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Name, SID, Disabled, Lockout FROM Win32_UserAccount WHERE LocalAccount=True")
WScript.Echo "Name" & vbTab & "SID" & vbTab & "Disabled" & vbTab & "Locked"
For Each account In users
    WScript.Echo account.Name & vbTab & account.SID & vbTab & _
        account.Disabled & vbTab & account.Lockout
Next
