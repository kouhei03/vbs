' ファイル概要: OSインストール日時。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 118_WindowsInstallDates.vbs
' 実行結果: OSインストール日時処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, systems, system, registry, result, installDate
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_OperatingSystem を照会し、InstallDate を取得します。
Set systems = service.ExecQuery("SELECT InstallDate FROM Win32_OperatingSystem")
For Each system In systems
    WScript.Echo "InstallDate=" & system.InstallDate
Next
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "InstallDate", installDate)
If result = 0 Then
    WScript.Echo "Setup InstallDate (Unix seconds)=" & installDate
Else
    WScript.Echo "Setup InstallDate unavailable; WMI code=" & result
End If
