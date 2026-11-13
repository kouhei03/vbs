' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: スクリーンセーバー関連設定と各照会結果を保持します。
Dim registry, result, value, valueName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 現在のユーザー設定からロック関連の値だけを読み取ります。
For Each valueName In Array("ScreenSaveActive", "ScreenSaveTimeOut", "ScreenSaverIsSecure", "SCRNSAVE.EXE")
    result = registry.GetStringValue(&H80000001, "Control Panel\Desktop", valueName, value)
    If result = 0 Then
        WScript.Echo valueName & "=" & value
    ElseIf result = 2 Then
        WScript.Echo valueName & "=(not set)"
    Else
        WScript.Echo valueName & " read failed; WMI return code: " & result
    End If
Next
