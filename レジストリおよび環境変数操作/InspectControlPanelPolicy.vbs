' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: コントロールパネル制限値を読み取るためのオブジェクトと結果を宣言します。
Dim registry, result, restricted
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: NoControlPanel値を読み、未設定と照会失敗を区別して表示します。
result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", restricted)
If result = 0 Then
    WScript.Echo "NoControlPanel=" & restricted
ElseIf result = 2 Then
    WScript.Echo "NoControlPanel is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
