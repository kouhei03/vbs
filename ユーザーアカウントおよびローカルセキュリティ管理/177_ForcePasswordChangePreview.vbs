Option Explicit

Dim userName, service, accounts, account, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 177_ForcePasswordChangePreview.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery( _
    "SELECT Name, PasswordChangeable, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
found = False
For Each account In accounts
    If StrComp(account.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Account: " & account.Name
        WScript.Echo "Password changeable: " & account.PasswordChangeable
        WScript.Echo "Password expires: " & account.PasswordExpires
    End If
Next
If Not found Then WScript.Echo "Local account was not found."
WScript.Echo "Preview only; no next-logon flag was set."
