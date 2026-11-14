' ローカルのプリンタードライバーとポートを一覧表示します。削除はしません。
Option Explicit
Dim service, drivers, driver, ports, port
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drivers = service.ExecQuery("SELECT Name, DriverPath, Version FROM Win32_PrinterDriver")
For Each driver In drivers
    WScript.Echo "Driver=" & driver.Name & " Version=" & driver.Version & " Path=" & driver.DriverPath
Next
Set ports = service.ExecQuery("SELECT Name, Description FROM Win32_TCPIPPrinterPort")
For Each port In ports
    WScript.Echo "TCPPort=" & port.Name & " " & port.Description
Next
