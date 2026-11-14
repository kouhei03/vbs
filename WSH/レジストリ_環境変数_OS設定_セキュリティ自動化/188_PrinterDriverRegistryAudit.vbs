' 概要: プリンタードライバー登録の件数を照会します。修復・削除は行いません。
' 実行方法: cscript //nologo 188_PrinterDriverRegistryAudit.vbs
' 処理結果: プリンタードライバー登録の読み取り監査。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, drivers, driver, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set drivers = service.ExecQuery("SELECT Name, DriverPath FROM Win32_PrinterDriver")
count = 0
For Each driver In drivers
    count = count + 1
    WScript.Echo "Driver=" & driver.Name & " Path=" & driver.DriverPath
Next
WScript.Echo "Driver registrations=" & count & " (read-only audit)."
