' ファイル概要: USBデバイスの名前・製造元・状態を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetUSBDevices.vbs
' 実行結果: USBデバイスの名前・製造元・状態を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, devices, device, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PnPEntity を照会し、Name, Manufacturer, Status を取得します。
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_PnPEntity " & _
    "WHERE PNPClass='USB'")

found = False
For Each device In devices
    found = True
    WScript.Echo "Device: " & device.Name
    WScript.Echo "Manufacturer: " & device.Manufacturer
    WScript.Echo "Status: " & device.Status
    WScript.Echo ""
Next
If Not found Then WScript.Echo "No USB devices returned by WMI."
