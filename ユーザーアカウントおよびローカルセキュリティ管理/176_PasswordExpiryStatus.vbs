Option Explicit

Dim service, accounts, account
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery( _
    "SELECT Name, PasswordExpires, PasswordRequired FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
WScript.Echo "Account" & vbTab & "PasswordExpires" & vbTab & "PasswordRequired"
For Each account In accounts
    WScript.Echo account.Name & vbTab & account.PasswordExpires & vbTab & _
        account.PasswordRequired
Next
WScript.Echo "Status only; password-expiry settings were not changed."
