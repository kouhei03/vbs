' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: レジストリプロバイダー、対象ハイブ、およびキー情報を保持します。
Dim registry, hiveName, hive, subKey, result, names, types
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo CheckRegistryExportSource.vbs <hive> <subkey>"
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
' 解説: エクスポート元キーにアクセスできるかだけを確認し、ファイル出力は行いません。
subKey = WScript.Arguments(1)
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.EnumValues(hive, subKey, names, types)
If result <> 0 Then
    WScript.Echo "Registry export source is unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Registry key is readable: " & hiveName & "\" & subKey
WScript.Echo "No .reg file was created; use an approved backup process if an export is required."
