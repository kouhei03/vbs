' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: WMIレジストリプロバイダー、入力値、取得結果を保持します。
Dim registry, hiveName, hive, subKey, valueName, valueType, result, data, item
If WScript.Arguments.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo InspectMultiStringAndBinary.vbs <hive> <subkey> <value-name> <MULTI_SZ|BINARY>"
    WScript.Quit 2
End If
' 解説: 対象ハイブをWMIの数値識別子へ対応付けます。
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
' 解説: 引数を格納し、要求された型に応じて読み取り専用メソッドを選びます。
subKey = WScript.Arguments(1)
valueName = WScript.Arguments(2)
valueType = UCase(WScript.Arguments(3))
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Select Case valueType
    Case "MULTI_SZ": result = registry.GetMultiStringValue(hive, subKey, valueName, data)
    Case "BINARY": result = registry.GetBinaryValue(hive, subKey, valueName, data)
    Case Else
        WScript.Echo "Value type must be MULTI_SZ or BINARY."
        WScript.Quit 2
End Select
' 解説: WMIの戻り値を確認し、読み取り失敗を成功扱いにしません。
If result <> 0 Then
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
' 解説: 複数行文字列は項目ごとに表示し、バイナリは内容を出さずサイズだけを表示します。
If valueType = "MULTI_SZ" Then
    If IsArray(data) Then
        For Each item In data
            WScript.Echo item
        Next
    End If
Else
    If IsArray(data) Then
        WScript.Echo "Binary byte count: " & (UBound(data) + 1)
    Else
        WScript.Echo "Binary byte count: 0"
    End If
End If
