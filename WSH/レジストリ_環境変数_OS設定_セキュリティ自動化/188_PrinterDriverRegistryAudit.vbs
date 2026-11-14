' プリンタードライバー登録の件数を照会します。修復・削除は行いません。
Option Explicit
Dim service, drivers, driver, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drivers = service.ExecQuery("SELECT Name, DriverPath FROM Win32_PrinterDriver")
count = 0
For Each driver In drivers
    count = count + 1
    WScript.Echo "Driver=" & driver.Name & " Path=" & driver.DriverPath
Next
WScript.Echo "Driver registrations=" & count & " (read-only audit)."
