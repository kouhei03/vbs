' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: 壁紙の現在値を読み取るためのレジストリプロバイダーと結果を保持します。
Dim registry, result, wallpaper
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 現在のユーザーの壁紙パスを読み取ります。強制設定や変更通知は行いません。
result = registry.GetStringValue(&H80000001, "Control Panel\Desktop", "Wallpaper", wallpaper)
If result <> 0 Then
    WScript.Echo "Wallpaper value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Wallpaper=" & wallpaper
