Option Explicit

Dim service, batteries, battery, found, batteryState
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set batteries = service.ExecQuery( _
    "SELECT EstimatedChargeRemaining, BatteryStatus, Status FROM Win32_Battery")

found = False
For Each battery In batteries
    found = True
    Select Case battery.BatteryStatus
        Case 1: batteryState = "Other"
        Case 2: batteryState = "Unknown"
        Case 3: batteryState = "Fully charged"
        Case 4: batteryState = "Low"
        Case 5: batteryState = "Critical"
        Case 6: batteryState = "Charging"
        Case 7: batteryState = "Charging and high"
        Case 8: batteryState = "Charging and low"
        Case 9: batteryState = "Charging and critical"
        Case 10: batteryState = "Undefined"
        Case 11: batteryState = "Partially charged"
        Case Else: batteryState = "Not reported"
    End Select
    WScript.Echo "Charge remaining (%): " & battery.EstimatedChargeRemaining
    WScript.Echo "Battery state: " & batteryState
    WScript.Echo "Health/status: " & battery.Status
Next
If Not found Then WScript.Echo "No battery reported (or battery data is unavailable)."
