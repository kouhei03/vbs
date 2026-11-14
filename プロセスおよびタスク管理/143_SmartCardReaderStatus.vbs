' ファイル概要: スマートカードリーダー状態照会。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 143_SmartCardReaderStatus.vbs
' 実行結果: スマートカードリーダー状態照会処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, devices, device
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set devices = service.ExecQuery("SELECT Name,PNPClass,Status,PNPDeviceID FROM Win32_PnPEntity WHERE PNPClass='SmartCardReader' OR Name LIKE '%smart card%'")
If devices.Count=0 Then WScript.Echo "No smart-card reader device was reported."
For Each device In devices
    WScript.Echo device.Name & " Status=" & device.Status & " PNPClass=" & device.PNPClass
Next
