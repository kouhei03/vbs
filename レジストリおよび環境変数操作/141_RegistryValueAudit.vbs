Option Explicit

Dim registry, hiveName, hive, subKey, result, valueNames, valueTypes
Dim subKeys, index, valueTypeName

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 141_RegistryValueAudit.vbs <hive> <subkey>"
    WScript.Quit 2
End If

hiveName = UCase(WScript.Arguments(0))
subKey = WScript.Arguments(1)
Select Case hiveName
    Case "HKCR": hive = &H80000000
    Case "HKCU": hive = &H80000001
    Case "HKLM": hive = &H80000002
    Case "HKU": hive = &H80000003
    Case "HKCC": hive = &H80000005
    Case Else
        WScript.Echo "Unsupported registry hive: " & hiveName
        WScript.Quit 2
End Select

Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
WScript.Echo "READ-ONLY registry inventory: " & hiveName & "\" & subKey

result = registry.EnumValues(hive, subKey, valueNames, valueTypes)
If result <> 0 Then
    WScript.Echo "Value enumeration failed with WMI return code: " & result
    WScript.Quit 1
End If

If IsArray(valueNames) Then
    For index = 0 To UBound(valueNames)
        Select Case valueTypes(index)
            Case 1: valueTypeName = "REG_SZ"
            Case 2: valueTypeName = "REG_EXPAND_SZ"
            Case 3: valueTypeName = "REG_BINARY"
            Case 4: valueTypeName = "REG_DWORD"
            Case 7: valueTypeName = "REG_MULTI_SZ"
            Case 11: valueTypeName = "REG_QWORD"
            Case Else: valueTypeName = "UNKNOWN(" & valueTypes(index) & ")"
        End Select
        WScript.Echo "Value name=[" & valueNames(index) & "] type=" & valueTypeName
    Next
Else
    WScript.Echo "No values found."
End If

result = registry.EnumKey(hive, subKey, subKeys)
If result <> 0 Then
    WScript.Echo "Subkey enumeration failed with WMI return code: " & result
    WScript.Quit 1
End If

If IsArray(subKeys) Then
    WScript.Echo "Immediate subkey count: " & (UBound(subKeys) + 1)
Else
    WScript.Echo "Immediate subkey count: 0"
End If
WScript.Echo "Value data was not read or displayed."
