' ファイル概要: UAC通知設定。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 132_UacNotificationSettings.vbs
' 実行結果: UAC通知設定処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, value, settingName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each settingName In Array("EnableLUA", "ConsentPromptBehaviorAdmin", "PromptOnSecureDesktop", "FilterAdministratorToken")
    ' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
    result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", settingName, value)
    If result = 0 Then
        WScript.Echo settingName & "=" & value
    Else
        WScript.Echo settingName & " unavailable; WMI code=" & result
    End If
Next
