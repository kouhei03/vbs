Option Explicit

Dim userName, service, accounts, account, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 173_DeleteLocalUserPreflight.vbs <user-name>"
    WScript.Quit 2
End If
userName = WScript.Arguments(0)
If Len(userName) = 0 Or Len(userName) > 256 Then
    WScript.Echo "User name length is invalid."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set accounts = service.ExecQuery( _
    "SELECT Name, SID, Disabled FROM Win32_UserAccount WHERE LocalAccount=True")
found = False
For Each account In accounts
    If StrComp(account.Name, userName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Candidate: " & account.Name
        WScript.Echo "SID: " & account.SID
        WScript.Echo "Disabled: " & account.Disabled
    End If
Next
If Not found Then WScript.Echo "Local account was not found."
WScript.Echo "Preflight only; no account or profile was deleted."
