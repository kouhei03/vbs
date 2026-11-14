' ファイル概要: 指定レジストリキーの直下のサブキーを列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListRegistrySubkeys.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定レジストリキーの直下のサブキーを列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim registry, hiveName, hive, subKey, result, children, child

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo ListRegistrySubkeys.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey>"
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
Set registry = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.EnumKey(hive, subKey, children)

If result <> 0 Then
    WScript.Echo "Registry enumeration failed with WMI return code: " & result
    WScript.Quit 1
End If

If IsArray(children) Then
    For Each child In children
        WScript.Echo child
    Next
Else
    WScript.Echo "No subkeys found."
End If
