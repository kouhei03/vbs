' ファイル概要: USBストレージの起動設定を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectUSBStoragePolicy.vbs
' 実行結果: USBストレージの起動設定を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, startValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Services\USBSTOR", "Start", startValue)
If result <> 0 Then
    WScript.Echo "USBSTOR Start value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "USBSTOR Start=" & startValue & " (read-only; no policy change made)"
