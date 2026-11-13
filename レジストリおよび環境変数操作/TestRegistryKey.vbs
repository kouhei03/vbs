' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（registry、hiveName、hive、subKey、parentKey、childKey）。主な用途: レジストリプロバイダー、レジストリハイブ名、レジストリハイブ識別子、レジストリサブキー。
Dim registry, hiveName, hive, subKey, parentKey, childKey
' 解説: このスクリプトで使う変数を宣言します（separator、result、children、child、found）。主な用途: 処理結果コード、対象が見つかったかを示すフラグ。
Dim separator, result, children, child, found

' 解説: 条件「WScript.Arguments.Count <> 2 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 2 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo TestRegistryKey.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey>"
    WScript.Echo "Usage: cscript //nologo TestRegistryKey.vbs <HKCU|HKLM|HKCR|HKU|HKCC> <subkey>"
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
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: separator = InStrRev(subKey, "\")
separator = InStrRev(subKey, "\")
' 解説: 条件「separator = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If separator = 0 Then
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: parentKey = ""
    parentKey = ""
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: childKey = subKey
    childKey = subKey
' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
Else
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: parentKey = Left(subKey, separator - 1)
    parentKey = Left(subKey, separator - 1)
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: childKey = Mid(subKey, separator + 1)
    childKey = Mid(subKey, separator + 1)
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 条件「Len(childKey) = 0 And Len(subKey) > 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(childKey) = 0 And Len(subKey) > 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Registry subkey must not end with a backslash."
    WScript.Echo "Registry subkey must not end with a backslash."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルWMIのStdRegProvプロバイダーへ接続し、レジストリ値の読み取りを行うオブジェクトを取得します。
Set registry = GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 指定したレジストリキー直下のサブキーを列挙し、戻り値コードで照会結果を確認します。
result = registry.EnumKey(hive, parentKey, children)

' 解説: 条件「result <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If result <> 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Could not enumerate the parent key; WMI return code: " & result
    WScript.Echo "Could not enumerate the parent key; WMI return code: " & result
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 条件「Len(subKey) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(subKey) = 0 Then
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
    found = True
' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
Else
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
    found = False
    ' 解説: 条件「IsArray(children) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If IsArray(children) Then
        ' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each child In children
        For Each child In children
            ' 解説: 条件「StrComp(child, childKey, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
            If StrComp(child, childKey, vbTextCompare) = 0 Then
                ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
                found = True
                ' 解説: 現在のForループを直ちに抜け、後続処理へ進みます。
                Exit For
            ' 解説: 条件分岐ブロックを終了します。
            End If
        ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
        Next
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 条件「found Then」を評価し、成立した場合に続く分岐処理を実行します。
If found Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Registry key exists: " & hiveName & "\" & subKey
    WScript.Echo "Registry key exists: " & hiveName & "\" & subKey
' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
Else
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Registry key not found: " & hiveName & "\" & subKey
    WScript.Echo "Registry key not found: " & hiveName & "\" & subKey
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
