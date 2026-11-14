' ファイル概要: OSインストール日時。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 118_WindowsInstallDates.vbs
' 実行結果: OSインストール日時処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, systems, system, registry, result, installDate
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT InstallDate FROM Win32_OperatingSystem")
For Each system In systems
    WScript.Echo "InstallDate=" & system.InstallDate
Next
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "InstallDate", installDate)
If result = 0 Then
    WScript.Echo "Setup InstallDate (Unix seconds)=" & installDate
Else
    WScript.Echo "Setup InstallDate unavailable; WMI code=" & result
End If
