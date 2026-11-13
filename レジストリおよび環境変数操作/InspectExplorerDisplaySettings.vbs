' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: エクスプローラーの表示値を読み取るための変数を保持します。
Dim registry, result, settingName, settingValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 隠しファイルと拡張子表示に関する現在のユーザー設定を読み取ります。
For Each settingName In Array("Hidden", "HideFileExt", "ShowSuperHidden")
    result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", settingName, settingValue)
    If result = 0 Then
        WScript.Echo settingName & "=" & settingValue
    ElseIf result = 2 Then
        WScript.Echo settingName & "=(not set)"
    Else
        WScript.Echo settingName & " read failed; WMI return code: " & result
    End If
Next
