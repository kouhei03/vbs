' ファイル概要: OSのサービスパックバージョンを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServicePack.vbs
' 実行結果: OSのサービスパックバージョンを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_OperatingSystem を照会し、ServicePackMajorVersion, ServicePackMinorVersion を取得します。
Set systems = service.ExecQuery( _
    "SELECT ServicePackMajorVersion, ServicePackMinorVersion FROM Win32_OperatingSystem")

For Each system In systems
    WScript.Echo "Service Pack: " & system.ServicePackMajorVersion & "." & _
        system.ServicePackMinorVersion
Next
