' 概要: Windowsバージョンとレジストリビュー情報を表示します。設定値は変更しません。
' 実行方法: cscript //nologo 195_WindowsCompatibilityAudit.vbs
' 処理結果: Windowsバージョン・アーキテクチャ確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item, shell, processor
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set systems = service.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
For Each item In systems
    WScript.Echo item.Caption & " Version=" & item.Version & " Build=" & item.BuildNumber
    WScript.Echo "Architecture=" & item.OSArchitecture
Next
Set shell = CreateObject("WScript.Shell")
processor = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
WScript.Echo "ProcessArchitecture=" & processor
WScript.Echo "Use documented OS-specific keys; no compatibility fallback was applied."
