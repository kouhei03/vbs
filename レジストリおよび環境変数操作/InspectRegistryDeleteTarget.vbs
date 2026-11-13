' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: レジストリプロバイダーと対象キーの列挙結果を保持します。
Dim registry, hiveName, hive, subKey, result, names, types, i
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo InspectRegistryDeleteTarget.vbs <hive> <subkey>"
    WScript.Quit 2
End If
' 解説: 許可済みハイブ名をWMI識別子へ変換します。
hiveName = UCase(WScript.Arguments(0))
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
' 解説: 指定キー内の値名だけを列挙し、値の削除は行いません。
subKey = WScript.Arguments(1)
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.EnumValues(hive, subKey, names, types)
If result <> 0 Then
    WScript.Echo "Could not enumerate key; WMI return code: " & result
    WScript.Quit 1
End If
' 解説: 対象キーと値名を報告し、削除の実行可能性を確認できるようにします。
WScript.Echo "READ-ONLY inspection; nothing was deleted."
WScript.Echo "Key: " & hiveName & "\" & subKey
If IsArray(names) Then
    For i = 0 To UBound(names)
        WScript.Echo "Value: " & names(i) & " (registry type " & types(i) & ")"
    Next
End If
