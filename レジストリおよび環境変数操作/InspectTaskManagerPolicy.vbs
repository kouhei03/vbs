' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: タスクマネージャーポリシー値を読み取るためのオブジェクトと結果を宣言します。
Dim registry, result, disabled
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 現在のユーザーのDisableTaskMgr値を読み、ポリシー状態を表示します。
result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", disabled)
If result = 0 Then
    WScript.Echo "DisableTaskMgr=" & disabled
ElseIf result = 2 Then
    WScript.Echo "DisableTaskMgr is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "No policy value was changed."
