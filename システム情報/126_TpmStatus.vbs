' 解説: TPM搭載/有効状態をローカルWMIから読み取り、値を表示します。システム設定は変更しません。
Option Explicit
' 解説: WMI接続、照会結果、対象インスタンス、およびプロパティ表示用変数を宣言します。
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
' 解説: 指定クラスを読み取り専用で照会します。
Set service = GetObject("winmgmts:\\.\root\cimv2\Security\MicrosoftTpm")
' 解説: 指定した情報を読み取り専用で照会します。
Set records = service.ExecQuery("SELECT SpecVersion, IsEnabled_InitialValue, IsActivated_InitialValue, ManufacturerIdTxt FROM Win32_Tpm")
' 解説: 対象がない場合は空結果であることを明示して終了します。
If records.Count = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "No data returned for: TPM搭載/有効状態"
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 0
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 各インスタンスの指定プロパティを、配列・NULLを考慮して表示します。
For Each record In records
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "---"
    ' 解説: コレクション内の各項目を順に処理します: For Each fieldName In Array("SpecVersion", "IsEnabled_InitialValue", "IsActivated_InitialValue", "ManufacturerIdTxt")
    For Each fieldName In Array("SpecVersion", "IsEnabled_InitialValue", "IsActivated_InitialValue", "ManufacturerIdTxt")
        ' 解説: 取得値または計算結果を変数へ格納します。
        fieldValue = record.Properties_.Item(fieldName).Value
        ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsNull(fieldValue)
        If IsNull(fieldValue) Then
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo fieldName & ": (not reported)"
        ' 解説: 先行条件が不成立の場合に追加条件を評価します: IsArray(fieldValue)
        ElseIf IsArray(fieldValue) Then
            ' 解説: 取得値または計算結果を変数へ格納します。
            joinedValue = ""
            ' 解説: コレクション内の各項目を順に処理します: For Each arrayItem In fieldValue
            For Each arrayItem In fieldValue
                ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Len(joinedValue) > 0
                If Len(joinedValue) > 0 Then joinedValue = joinedValue & ", "
                ' 解説: 取得値または計算結果を変数へ格納します。
                joinedValue = joinedValue & CStr(arrayItem)
            ' 解説: 現在の反復を終え、次の要素へ進みます。
            Next
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo fieldName & ": " & joinedValue
        ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
        Else
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo fieldName & ": " & CStr(fieldValue)
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
