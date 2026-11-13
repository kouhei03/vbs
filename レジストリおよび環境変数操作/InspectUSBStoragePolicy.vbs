' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: USBSTORの起動設定を読むためのプロバイダーと結果変数を用意します。
Dim registry, result, startValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: USBストレージドライバーのStart値を読み取ります。設定変更は行いません。
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Services\USBSTOR", "Start", startValue)
If result <> 0 Then
    WScript.Echo "USBSTOR Start value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "USBSTOR Start=" & startValue & " (read-only; no policy change made)"
