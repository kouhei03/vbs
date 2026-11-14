' ファイル概要: OSのアーキテクチャを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetOSArchitecture.vbs
' 実行結果: OSのアーキテクチャを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_OperatingSystem を照会し、OSArchitecture を取得します。
Set systems = service.ExecQuery("SELECT OSArchitecture FROM Win32_OperatingSystem")

For Each system In systems
    WScript.Echo "OS architecture: " & system.OSArchitecture
Next
