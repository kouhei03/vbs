' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: UAC関連DWORD値と照会結果を保持します。
Dim registry, result, settingName, settingValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: UACの主要な有効化・昇格動作設定を読み取り、値の書き換えは行いません。
For Each settingName In Array("EnableLUA", "ConsentPromptBehaviorAdmin", "PromptOnSecureDesktop", "FilterAdministratorToken")
    result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", settingName, settingValue)
    If result = 0 Then
        WScript.Echo settingName & "=" & settingValue
    ElseIf result = 2 Then
        WScript.Echo settingName & "=(not set)"
    Else
        WScript.Echo settingName & " read failed; WMI return code: " & result
    End If
Next
