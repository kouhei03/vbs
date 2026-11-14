Option Explicit

Dim service, settings, setting, shareName, descriptor, result
Dim entry, trustee, trusteeName, accessType
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetShareAcl.vbs <share-name>"
    WScript.Quit 2
End If
shareName = Trim(WScript.Arguments(0))
If Len(shareName) = 0 Then
    WScript.Echo "Share name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set settings = service.ExecQuery( _
    "SELECT Name FROM Win32_LogicalShareSecuritySetting")
For Each setting In settings
    If StrComp(setting.Name, shareName, vbTextCompare) = 0 Then
        result = setting.GetSecurityDescriptor(descriptor)
        If result <> 0 Then
            WScript.Echo "Could not read share ACL; return code: " & result
            WScript.Quit 1
        End If
        WScript.Echo "Share: " & setting.Name
        If IsArray(descriptor.DACL) Then
            WScript.Echo "Trustee" & vbTab & "ACE type" & vbTab & "Access mask"
            For Each entry In descriptor.DACL
                Set trustee = entry.Trustee
                trusteeName = trustee.Name
                If Not IsNull(trustee.Domain) Then
                    If Len(trustee.Domain) > 0 Then
                        trusteeName = trustee.Domain & "\" & trusteeName
                    End If
                End If
                Select Case entry.AceType
                    Case 0: accessType = "Allow"
                    Case 1: accessType = "Deny"
                    Case Else: accessType = "Other (" & entry.AceType & ")"
                End Select
                WScript.Echo trusteeName & vbTab & accessType & vbTab & _
                    entry.AccessMask
            Next
        Else
            WScript.Echo "No DACL entries returned."
        End If
        WScript.Quit 0
    End If
Next

WScript.Echo "Share ACL not found: " & shareName
WScript.Quit 3
