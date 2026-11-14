Option Explicit

Dim service, groups, group
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set groups = service.ExecQuery( _
    "SELECT Name, Description FROM Win32_Group WHERE LocalAccount=True")

WScript.Echo "Local group" & vbTab & "Description"
For Each group In groups
    WScript.Echo group.Name & vbTab & group.Description
Next
