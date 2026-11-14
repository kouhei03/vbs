' 概要: OS、UAC、環境変数の構成状態をまとめて表示します。設定変更はしません。
' 実行方法: cscript //nologo 200_ReadOnlyEnvironmentLauncher.vbs
' 処理結果: OS・環境・UACの読み取り専用サマリー。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item, shell, env, value, keyPath
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set systems = service.ExecQuery("SELECT Caption, Version, BuildNumber FROM Win32_OperatingSystem")
For Each item In systems
    WScript.Echo "OS=" & item.Caption & " Version=" & item.Version & " Build=" & item.BuildNumber
Next
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
WScript.Echo "PROCESSOR_ARCHITECTURE=" & env("PROCESSOR_ARCHITECTURE")
WScript.Echo "TEMP configured=" & CStr(env("TEMP") <> "")
keyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA"
On Error Resume Next
value = shell.RegRead(keyPath)
If Err.Number = 0 Then
    WScript.Echo "UAC EnableLUA=" & value
Else
    WScript.Echo "UAC status unavailable: " & Err.Description
End If
On Error GoTo 0
WScript.Echo "Read-only summary; no registry, environment, or security settings changed."
