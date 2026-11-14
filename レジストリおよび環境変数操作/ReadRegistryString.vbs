Option Explicit

Dim registry, hiveName, hive, subKey, valueName, value, result

If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo ReadRegistryString.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey> <value-name>"
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
valueName = WScript.Arguments(2)
Set registry = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(hive, subKey, valueName, value)

If result <> 0 Then
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If

If IsNull(value) Then value = ""
WScript.Echo hiveName & "\" & subKey & "\" & valueName & " = " & value
