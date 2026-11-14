Option Explicit

Dim service, settings, setting, descriptor, result, entry, trustee
Dim domainName, trusteeName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set settings = service.ExecQuery( _
    "SELECT Name FROM Win32_LogicalShareSecuritySetting")
For Each setting In settings
    result = setting.GetSecurityDescriptor(descriptor)
    If result = 0 Then
        WScript.Echo "Share: " & setting.Name
        If IsArray(descriptor.DACL) Then
            For Each entry In descriptor.DACL
                Set trustee = entry.Trustee
                domainName = ""
                trusteeName = ""
                If Not IsNull(trustee.Domain) Then domainName = CStr(trustee.Domain)
                If Not IsNull(trustee.Name) Then trusteeName = CStr(trustee.Name)
                If Len(domainName) > 0 Then trusteeName = domainName & "\" & trusteeName
                WScript.Echo "  " & trusteeName & " ACE=" & entry.AceType & _
                    " mask=" & entry.AccessMask
            Next
        End If
    Else
        WScript.Echo "Share ACL unavailable: " & setting.Name & _
            " (return code " & result & ")"
    End If
Next
WScript.Echo "ACLs were inspected only; no permissions were changed."
