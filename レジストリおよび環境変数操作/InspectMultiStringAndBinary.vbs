' ファイル概要: 複数行文字列またはバイナリ値のサイズを読み取る。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectMultiStringAndBinary.vbs <hive> <subkey> <value-name> <MULTI_SZ|BINARY>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 複数行文字列またはバイナリ値のサイズを読み取る処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, hiveName, hive, subKey, valueName, valueType, result, data, item
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo InspectMultiStringAndBinary.vbs <hive> <subkey> <value-name> <MULTI_SZ|BINARY>"
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
valueType = UCase(WScript.Arguments(3))
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Select Case valueType
    ' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
    Case "MULTI_SZ": result = registry.GetMultiStringValue(hive, subKey, valueName, data)
    Case "BINARY": result = registry.GetBinaryValue(hive, subKey, valueName, data)
    Case Else
        WScript.Echo "Value type must be MULTI_SZ or BINARY."
        WScript.Quit 2
End Select
If result <> 0 Then
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
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
