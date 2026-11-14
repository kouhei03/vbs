Option Explicit

Dim service, shares, share
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set shares = service.ExecQuery( _
    "SELECT Name, Path, Type, Description FROM Win32_Share")
WScript.Echo "Name" & vbTab & "Path" & vbTab & "Type" & vbTab & "Description"
For Each share In shares
    WScript.Echo share.Name & vbTab & share.Path & vbTab & _
        share.Type & vbTab & share.Description
Next
