Option Explicit

Dim service, groups, group, members, member, domainName, memberName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set groups = service.ExecQuery( _
    "SELECT Name, SID FROM Win32_Group WHERE SID='S-1-5-32-544'")
For Each group In groups
    WScript.Echo "Local Administrators group: " & group.Name
    Set members = group.Associators_("Win32_GroupUser")
    For Each member In members
        domainName = ""
        memberName = ""
        If Not IsNull(member.Domain) Then domainName = CStr(member.Domain)
        If Not IsNull(member.Name) Then memberName = CStr(member.Name)
        If Len(domainName) > 0 Then memberName = domainName & "\" & memberName
        WScript.Echo memberName
    Next
Next
WScript.Echo "Membership was inspected only; no user was added."
