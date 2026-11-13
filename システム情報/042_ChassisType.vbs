' 解説: SMBIOS筐体種別コードを読み取り、代表的なコードを人が読める名称に変換します。
Option Explicit
' 解説: WMI接続、筐体情報、および配列要素を保持します。
Dim service, enclosures, enclosure, code, label
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set enclosures = service.ExecQuery("SELECT ChassisTypes FROM Win32_SystemEnclosure")
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: enclosures.Count = 0
If enclosures.Count = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "No chassis information returned."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: コレクション内の各項目を順に処理します: For Each enclosure In enclosures
For Each enclosure In enclosures
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsArray(enclosure.ChassisTypes)
    If IsArray(enclosure.ChassisTypes) Then
        ' 解説: コレクション内の各項目を順に処理します: For Each code In enclosure.ChassisTypes
        For Each code In enclosure.ChassisTypes
            ' 解説: 値に応じて対応するケースを選択します: CLng(code)
            Select Case CLng(code)
                ' 解説: この値または条件に該当する場合の分岐です: Case 3: label = "Desktop"
                Case 3: label = "Desktop"
                ' 解説: この値または条件に該当する場合の分岐です: Case 8, 9, 10, 14: label = "Portable/Laptop/Notebook"
                Case 8, 9, 10, 14: label = "Portable/Laptop/Notebook"
                ' 解説: この値または条件に該当する場合の分岐です: Case 11: label = "Handheld"
                Case 11: label = "Handheld"
                ' 解説: この値または条件に該当する場合の分岐です: Case 13: label = "All-in-One"
                Case 13: label = "All-in-One"
                ' 解説: この値または条件に該当する場合の分岐です: Case 30: label = "Tablet"
                Case 30: label = "Tablet"
                ' 解説: この値または条件に該当する場合の分岐です: Case 31: label = "Convertible"
                Case 31: label = "Convertible"
                ' 解説: この値または条件に該当する場合の分岐です: Case 32: label = "Detachable"
                Case 32: label = "Detachable"
                ' 解説: この値または条件に該当する場合の分岐です: Case 7: label = "Tower"
                Case 7: label = "Tower"
                ' 解説: この値または条件に該当する場合の分岐です: Case 23: label = "Rack Mount"
                Case 23: label = "Rack Mount"
                ' 解説: この値または条件に該当する場合の分岐です: Case Else: label = "Other SMBIOS chassis type"
                Case Else: label = "Other SMBIOS chassis type"
            ' 解説: Select Caseの選択分岐を終了します。
            End Select
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo "Chassis type code " & code & ": " & label
        ' 解説: 現在の反復を終え、次の要素へ進みます。
        Next
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Chassis type data is unavailable."
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
