Option Explicit

Dim groupName, userName, service, groups, group, members, member
Dim groupFound, memberFound, memberDomain, memberName, separator, memberLabel
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 179_GroupMemberRemovalPreview.vbs <local-group> <member>"
    WScript.Quit 2
End If
groupName = WScript.Arguments(0)
userName = WScript.Arguments(1)
If Not IsSafeName(groupName) Or Not IsSafeName(userName) Then
    WScript.Echo "Names may contain only letters, digits, dot, underscore, hyphen, and backslash."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set groups = service.ExecQuery( _
    "SELECT Name FROM Win32_Group WHERE LocalAccount=True")
groupFound = False
memberFound = False
For Each group In groups
    If StrComp(group.Name, groupName, vbTextCompare) = 0 Then
        groupFound = True
        WScript.Echo "Local group: " & group.Name
        Set members = group.Associators_("Win32_GroupUser")
        For Each member In members
            memberName = ""
            memberDomain = ""
            If Not IsNull(member.Name) Then memberName = CStr(member.Name)
            If Not IsNull(member.Domain) Then memberDomain = member.Domain
            memberLabel = memberName
            If Len(memberDomain) > 0 Then memberLabel = memberDomain & "\" & memberName
            WScript.Echo "Member: " & memberLabel
            separator = InStr(1, userName, "\", vbBinaryCompare)
            If separator > 0 Then
                If StrComp(memberLabel, userName, vbTextCompare) = 0 Then
                    memberFound = True
                End If
            ElseIf StrComp(memberName, userName, vbTextCompare) = 0 Then
                memberFound = True
            End If
        Next
    End If
Next
If Not groupFound Then
    WScript.Echo "Local group was not found."
ElseIf memberFound Then
    WScript.Echo "Requested member is currently present."
Else
    WScript.Echo "Requested member was not found in the group."
End If
WScript.Echo "Preflight only; no group membership was changed."

Function IsSafeName(value)
    Dim expression
    Set expression = New RegExp
    expression.Pattern = "^[A-Za-z0-9._-]+(\\[A-Za-z0-9._-]+)?$"
    IsSafeName = expression.Test(value)
End Function
