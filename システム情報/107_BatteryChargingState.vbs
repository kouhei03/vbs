Option Explicit
Dim service, batteries, battery, stateName
Set service = GetObject("winmgmts:\\.\root\cimv2")
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
