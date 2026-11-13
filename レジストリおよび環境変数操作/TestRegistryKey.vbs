Option Explicit

Dim registry, hiveName, hive, subKey, parentKey, childKey
Dim separator, result, children, child, found

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo TestRegistryKey.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey>"
    WScript.Quit 2
End If

hiveName = UCase(WScript.Arguments(0))
Select Case hiveName
    Case "HKCU": hive = &H80000001
    Case "HKLM": hive = &H80000002
    Case "HKCR": hive = &H80000000
    Case "HKU": hive = &H80000003
    Case "HKCC": hive = &H80000005
    Case Else
        WScript.Echo "Unsupported registry hive: " & hiveName
        WScript.Quit 2
End Select

subKey = WScript.Arguments(1)
separator = InStrRev(subKey, "\")
If separator = 0 Then
    parentKey = ""
    childKey = subKey
Else
    parentKey = Left(subKey, separator - 1)
    childKey = Mid(subKey, separator + 1)
End If

If Len(childKey) = 0 And Len(subKey) > 0 Then
    WScript.Echo "Registry subkey must not end with a backslash."
    WScript.Quit 2
End If

Set registry = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.EnumKey(hive, parentKey, children)

If result <> 0 Then
    WScript.Echo "Could not enumerate the parent key; WMI return code: " & result
    WScript.Quit 1
End If

If Len(subKey) = 0 Then
    found = True
Else
    found = False
    If IsArray(children) Then
        For Each child In children
            If StrComp(child, childKey, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next
    End If
End If

If found Then
    WScript.Echo "Registry key exists: " & hiveName & "\" & subKey
Else
    WScript.Echo "Registry key not found: " & hiveName & "\" & subKey
    WScript.Quit 1
End If
