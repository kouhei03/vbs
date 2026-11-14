' OS、UAC、環境変数の構成状態をまとめて表示します。設定変更はしません。
Option Explicit
Dim service, systems, item, shell, env, value, keyPath
Set service = GetObject("winmgmts:\\.\root\cimv2")
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
