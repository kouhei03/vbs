' Windowsバージョンとレジストリビュー情報を表示します。設定値は変更しません。
Option Explicit
Dim service, systems, item, shell, processor
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
For Each item In systems
    WScript.Echo item.Caption & " Version=" & item.Version & " Build=" & item.BuildNumber
    WScript.Echo "Architecture=" & item.OSArchitecture
Next
Set shell = CreateObject("WScript.Shell")
processor = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
WScript.Echo "ProcessArchitecture=" & processor
WScript.Echo "Use documented OS-specific keys; no compatibility fallback was applied."
