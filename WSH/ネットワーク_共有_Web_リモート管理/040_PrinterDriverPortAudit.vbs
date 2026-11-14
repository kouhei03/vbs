' 概要: ローカルのプリンタードライバーとポートを一覧表示します。削除はしません。
' 実行方法: cscript //nologo 040_PrinterDriverPortAudit.vbs
' 処理結果: プリンタードライバー・TCPポート一覧。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, drivers, driver, ports, port
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set drivers = service.ExecQuery("SELECT Name, DriverPath, Version FROM Win32_PrinterDriver")
For Each driver In drivers
    WScript.Echo "Driver=" & driver.Name & " Version=" & driver.Version & " Path=" & driver.DriverPath
Next
Set ports = service.ExecQuery("SELECT Name, Description FROM Win32_TCPIPPrinterPort")
For Each port In ports
    WScript.Echo "TCPPort=" & port.Name & " " & port.Description
Next
