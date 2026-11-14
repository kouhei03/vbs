Option Explicit
Dim args, hiveName, subKey, valueName, proposedValue
Set args = WScript.Arguments
If args.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo PreviewRegistryStringWrite.vbs <hive> <subkey> <value-name> <value>"
    WScript.Quit 2
End If
hiveName = UCase(args(0))
subKey = args(1)
valueName = args(2)
proposedValue = args(3)
If InStr("|HKCU|HKLM|HKCR|HKU|HKCC|", "|" & hiveName & "|") = 0 Then
    WScript.Echo "Unsupported registry hive: " & hiveName
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no registry value was written."
WScript.Echo "Target: " & hiveName & "\" & subKey & "\" & valueName
WScript.Echo "Proposed REG_SZ character count: " & Len(proposedValue)
