' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（registry、hiveName、hive、subKey、valueName、value、result）。主な用途: レジストリプロバイダー、レジストリハイブ名、レジストリハイブ識別子、レジストリサブキー、レジストリ値名、取得した値、処理結果コード。
Dim registry, hiveName, hive, subKey, valueName, value, result

' 解説: 条件「WScript.Arguments.Count <> 3 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 3 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo ReadRegistryString.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey> <value-name>"
    WScript.Echo "Usage: cscript //nologo ReadRegistryString.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey> <value-name>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: hiveName = UCase(WScript.Arguments(0))
hiveName = UCase(WScript.Arguments(0))
' 解説: 値ごとに処理を分ける選択分岐を開始します: hiveName
Select Case hiveName
    ' 解説: Select Caseの値が「"HKCU": hive = &H80000001」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
    Case "HKCU": hive = &H80000001
    ' 解説: Select Caseの値が「"HKLM": hive = &H80000002」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
    Case "HKLM": hive = &H80000002
    ' 解説: Select Caseの値が「"HKCR": hive = &H80000000」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
    Case "HKCR": hive = &H80000000
    ' 解説: Select Caseの値が「"HKU": hive = &H80000003」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
    Case "HKU": hive = &H80000003
    ' 解説: Select Caseの値が「"HKCC": hive = &H80000005」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
    Case "HKCC": hive = &H80000005
    ' 解説: Select CaseのどのCase値にも一致しない場合に実行する既定分岐です。
    Case Else
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Unsupported registry hive: " & hiveName
        WScript.Echo "Unsupported registry hive: " & hiveName
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
        WScript.Quit 2
' 解説: Select Caseの選択分岐を終了します。
End Select

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: subKey = WScript.Arguments(1)
subKey = WScript.Arguments(1)
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: valueName = WScript.Arguments(2)
valueName = WScript.Arguments(2)
' 解説: ローカルWMIのStdRegProvプロバイダーへ接続し、レジストリ値の読み取りを行うオブジェクトを取得します。
Set registry = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 指定したレジストリハイブ/キーからREG_SZ文字列値を読み取ります。レジストリは変更しません。
result = registry.GetStringValue(hive, subKey, valueName, value)

' 解説: 条件「result <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If result <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Registry read failed with WMI return code: " & result
    WScript.Echo "Registry read failed with WMI return code: " & result
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 条件「IsNull(value) Then value = ""」を評価し、成立した場合に続く分岐処理を実行します。
If IsNull(value) Then value = ""
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: hiveName & "\" & subKey & "\" & valueName & " = " & value
WScript.Echo hiveName & "\" & subKey & "\" & valueName & " = " & value
