' ファイル概要: 削除候補キーの値名と型を確認する（削除なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectRegistryDeleteTarget.vbs <hive> <subkey>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 削除候補キーの値名と型を確認する（削除なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, hiveName, hive, subKey, result, names, types, i
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo InspectRegistryDeleteTarget.vbs <hive> <subkey>"
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
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定キーの値名と型を列挙します。値データはこの操作では取得しません。
result = registry.EnumValues(hive, subKey, names, types)
If result <> 0 Then
    WScript.Echo "Could not enumerate key; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "READ-ONLY inspection; nothing was deleted."
WScript.Echo "Key: " & hiveName & "\" & subKey
If IsArray(names) Then
    For i = 0 To UBound(names)
        WScript.Echo "Value: " & names(i) & " (registry type " & types(i) & ")"
    Next
End If
