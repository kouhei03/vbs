Option Explicit
Dim service, designRows, fullRows, row, foundDesign, foundFull
Set service = GetObject("winmgmts:\\.\root\wmi")
Set designRows = service.ExecQuery("SELECT DesignedCapacity FROM BatteryStaticData")
Set fullRows = service.ExecQuery("SELECT FullChargedCapacity FROM BatteryFullChargedCapacity")
foundDesign = False
foundFull = False
For Each row In designRows
    WScript.Echo "DesignedCapacity=" & row.DesignedCapacity
    foundDesign = True
Next
For Each row In fullRows
    WScript.Echo "FullChargedCapacity=" & row.FullChargedCapacity
    foundFull = True
Next
If Not foundDesign Or Not foundFull Then
    WScript.Echo "Battery capacity data is not available from this device/provider."
    WScript.Quit 1
End If
