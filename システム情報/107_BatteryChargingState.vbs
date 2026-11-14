' ファイル概要: バッテリー充電状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 107_BatteryChargingState.vbs
' 実行結果: バッテリー充電状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, batteries, battery, stateName
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Battery を照会し、DeviceID, BatteryStatus, EstimatedChargeRemaining, Status を取得します。
Set batteries = service.ExecQuery("SELECT DeviceID, BatteryStatus, EstimatedChargeRemaining, Status FROM Win32_Battery")
If batteries.Count = 0 Then
    WScript.Echo "No battery reported by WMI."
    WScript.Quit 0
End If
For Each battery In batteries
    Select Case battery.BatteryStatus
        Case 1: stateName = "Other"
        Case 2: stateName = "Unknown"
        Case 3: stateName = "Fully charged"
        Case 4: stateName = "Low"
        Case 5: stateName = "Critical"
        Case 6: stateName = "Charging"
        Case 7: stateName = "Charging and high"
        Case 8: stateName = "Charging and low"
        Case 9: stateName = "Charging and critical"
        Case 10: stateName = "Undefined"
        Case 11: stateName = "Partially charged"
        Case Else: stateName = "Unreported status " & battery.BatteryStatus
    End Select
    WScript.Echo battery.DeviceID & ": " & stateName
    WScript.Echo "Remaining percent=" & battery.EstimatedChargeRemaining & " WMI status=" & battery.Status
Next
