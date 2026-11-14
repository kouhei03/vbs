' ファイル概要: キーボードとポインティングデバイスを列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetInputDevices.vbs
' 実行結果: キーボードとポインティングデバイスを列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, devices, device
Set service = GetObject("winmgmts:\\.\root\cimv2")

WScript.Echo "Keyboards:"
' [処理解説] WMIの Win32_Keyboard を照会し、Description を取得します。
Set devices = service.ExecQuery("SELECT Description FROM Win32_Keyboard")
For Each device In devices
    WScript.Echo "  " & device.Description
Next

WScript.Echo "Pointing devices:"
' [処理解説] WMIの Win32_PointingDevice を照会し、Description を取得します。
Set devices = service.ExecQuery("SELECT Description FROM Win32_PointingDevice")
For Each device In devices
    WScript.Echo "  " & device.Description
Next
