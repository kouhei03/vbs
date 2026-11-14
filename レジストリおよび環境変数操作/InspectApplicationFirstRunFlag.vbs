' ファイル概要: 指定したレジストリ値の存在を確認する（変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectApplicationFirstRunFlag.vbs <hive> <subkey> <value-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定したレジストリ値の存在を確認する（変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, hiveName, hive, subKey, valueName, result, valueNames, valueTypes, i
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo InspectApplicationFirstRunFlag.vbs <hive> <subkey> <value-name>"
    WScript.Quit 2
End If
hiveName = UCase(WScript.Arguments(0))
' [処理解説] 入力されたハイブ名をStdRegProvで使用する定数に対応付けます。
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
subKey = WScript.Arguments(1)
valueName = WScript.Arguments(2)
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定キーの値名と型を列挙します。値データはこの操作では取得しません。
result = registry.EnumValues(hive, subKey, valueNames, valueTypes)
If result <> 0 Then
    WScript.Echo "Could not enumerate key; WMI return code: " & result
    WScript.Quit 1
End If
If IsArray(valueNames) Then
    For i = 0 To UBound(valueNames)
        If StrComp(valueNames(i), valueName, vbTextCompare) = 0 Then
            WScript.Echo "Value exists; registry type=" & valueTypes(i)
            WScript.Echo "Use ReadRegistryDword.vbs or ReadRegistryString.vbs to inspect its data."
            WScript.Quit 0
        End If
    Next
End If
WScript.Echo "Value is not present."
