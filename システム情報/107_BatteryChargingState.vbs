' 解説: バッテリー状態コードと残量を読み取り、WMIの代表的な状態値を名称に変換します。
Option Explicit
' 解説: WMI接続、バッテリー情報、および状態表示名を保持します。
Dim service, batteries, battery, stateName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set batteries = service.ExecQuery("SELECT DeviceID, BatteryStatus, EstimatedChargeRemaining, Status FROM Win32_Battery")
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: batteries.Count = 0
If batteries.Count = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "No battery reported by WMI."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 0
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: コレクション内の各項目を順に処理します: For Each battery In batteries
For Each battery In batteries
    ' 解説: 値に応じて対応するケースを選択します: battery.BatteryStatus
    Select Case battery.BatteryStatus
        ' 解説: この値または条件に該当する場合の分岐です: Case 1: stateName = "Other"
        Case 1: stateName = "Other"
        ' 解説: この値または条件に該当する場合の分岐です: Case 2: stateName = "Unknown"
        Case 2: stateName = "Unknown"
        ' 解説: この値または条件に該当する場合の分岐です: Case 3: stateName = "Fully charged"
        Case 3: stateName = "Fully charged"
        ' 解説: この値または条件に該当する場合の分岐です: Case 4: stateName = "Low"
        Case 4: stateName = "Low"
        ' 解説: この値または条件に該当する場合の分岐です: Case 5: stateName = "Critical"
        Case 5: stateName = "Critical"
        ' 解説: この値または条件に該当する場合の分岐です: Case 6: stateName = "Charging"
        Case 6: stateName = "Charging"
        ' 解説: この値または条件に該当する場合の分岐です: Case 7: stateName = "Charging and high"
        Case 7: stateName = "Charging and high"
        ' 解説: この値または条件に該当する場合の分岐です: Case 8: stateName = "Charging and low"
        Case 8: stateName = "Charging and low"
        ' 解説: この値または条件に該当する場合の分岐です: Case 9: stateName = "Charging and critical"
        Case 9: stateName = "Charging and critical"
        ' 解説: この値または条件に該当する場合の分岐です: Case 10: stateName = "Undefined"
        Case 10: stateName = "Undefined"
        ' 解説: この値または条件に該当する場合の分岐です: Case 11: stateName = "Partially charged"
        Case 11: stateName = "Partially charged"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: stateName = "Unreported status " & battery.BatteryStatus
        Case Else: stateName = "Unreported status " & battery.BatteryStatus
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo battery.DeviceID & ": " & stateName
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Remaining percent=" & battery.EstimatedChargeRemaining & " WMI status=" & battery.Status
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
