' ファイル概要: オーディオデバイスの情報を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetAudioDevices.vbs
' 実行結果: オーディオデバイスの情報を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, devices, device
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_SoundDevice")

For Each device In devices
    WScript.Echo "Device: " & device.Name
    WScript.Echo "Manufacturer: " & device.Manufacturer
    WScript.Echo "Status: " & device.Status
    WScript.Echo ""
Next
